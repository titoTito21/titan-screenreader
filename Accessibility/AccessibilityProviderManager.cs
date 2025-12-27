using System.Windows.Automation;

namespace ScreenReader.Accessibility;

/// <summary>
/// Zarządza dostawcami technologii dostępnościowych
/// Port z NVDA - pozwala na przełączanie między różnymi API
/// </summary>
public class AccessibilityProviderManager : IDisposable
{
    private readonly Dictionary<AccessibilityAPI, IAccessibilityProvider> _providers = new();
    private AccessibilityAPI _enabledApis = AccessibilityAPI.UIAutomation;
    private AccessibilityAPI _preferredApi = AccessibilityAPI.UIAutomation;
    private bool _disposed;

    /// <summary>
    /// Aktualnie włączone API
    /// </summary>
    public AccessibilityAPI EnabledApis => _enabledApis;

    /// <summary>
    /// Preferowane API (używane jako pierwsze)
    /// </summary>
    public AccessibilityAPI PreferredApi
    {
        get => _preferredApi;
        set
        {
            _preferredApi = value;
            Console.WriteLine($"AccessibilityProviderManager: Preferowane API zmienione na {value.GetPolishName()}");
        }
    }

    /// <summary>
    /// Event wywoływany przy zmianie fokusu
    /// </summary>
    public event EventHandler<AccessibleObject>? FocusChanged;

    /// <summary>
    /// Event wywoływany przy zmianie API
    /// </summary>
    public event EventHandler<AccessibilityAPI>? ApiChanged;

    public AccessibilityProviderManager()
    {
        RegisterDefaultProviders();
    }

    /// <summary>
    /// Rejestruje domyślnych providerów
    /// </summary>
    private void RegisterDefaultProviders()
    {
        // UI Automation - zawsze dostępne
        RegisterProvider(new UIAutomationProvider());

        // MSAA
        RegisterProvider(new MSAAProvider());

        // IAccessible2 (dla Firefox, LibreOffice)
        RegisterProvider(new IAccessible2Provider());

        // Java Access Bridge
        RegisterProvider(new JavaAccessBridgeProvider());

        Console.WriteLine($"AccessibilityProviderManager: Zarejestrowano {_providers.Count} providerów");
    }

    /// <summary>
    /// Rejestruje providera
    /// </summary>
    public void RegisterProvider(IAccessibilityProvider provider)
    {
        if (!_providers.ContainsKey(provider.ApiType))
        {
            _providers[provider.ApiType] = provider;

            // Podłącz eventy
            provider.FocusChanged += OnProviderFocusChanged;

            Console.WriteLine($"AccessibilityProviderManager: Zarejestrowano provider {provider.ApiType.GetPolishName()}");
        }
    }

    /// <summary>
    /// Inicjalizuje wszystkich włączonych providerów
    /// </summary>
    public void Initialize()
    {
        foreach (var kvp in _providers)
        {
            if (_enabledApis.HasFlag(kvp.Key))
            {
                try
                {
                    if (kvp.Value.IsAvailable && kvp.Value.Initialize())
                    {
                        Console.WriteLine($"AccessibilityProviderManager: Zainicjalizowano {kvp.Key.GetPolishName()}");
                    }
                    else
                    {
                        Console.WriteLine($"AccessibilityProviderManager: {kvp.Key.GetPolishName()} niedostępny");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"AccessibilityProviderManager: Błąd inicjalizacji {kvp.Key.GetPolishName()}: {ex.Message}");
                }
            }
        }
    }

    /// <summary>
    /// Włącza/wyłącza API
    /// </summary>
    public void SetApiEnabled(AccessibilityAPI api, bool enabled)
    {
        if (enabled)
        {
            _enabledApis |= api;
            if (_providers.TryGetValue(api, out var provider) && provider.IsAvailable)
            {
                provider.Initialize();
                provider.StartEventListening();
            }
        }
        else
        {
            _enabledApis &= ~api;
            if (_providers.TryGetValue(api, out var provider))
            {
                provider.StopEventListening();
            }
        }

        Console.WriteLine($"AccessibilityProviderManager: {api.GetPolishName()} {(enabled ? "włączone" : "wyłączone")}");
        ApiChanged?.Invoke(this, _enabledApis);
    }

    /// <summary>
    /// Przełącza API na następne dostępne
    /// </summary>
    public string CyclePreferredApi()
    {
        var apis = new[] { AccessibilityAPI.UIAutomation, AccessibilityAPI.MSAA, AccessibilityAPI.IAccessible2, AccessibilityAPI.JavaAccessBridge };
        int currentIndex = Array.IndexOf(apis, _preferredApi);
        int nextIndex = (currentIndex + 1) % apis.Length;

        // Znajdź następne dostępne API
        for (int i = 0; i < apis.Length; i++)
        {
            int idx = (nextIndex + i) % apis.Length;
            if (_providers.TryGetValue(apis[idx], out var provider) && provider.IsAvailable)
            {
                PreferredApi = apis[idx];
                return $"Preferowane API: {PreferredApi.GetPolishName()}";
            }
        }

        return "Brak dostępnych API";
    }

    /// <summary>
    /// Pobiera obiekt z fokusem używając najlepszego dostępnego API
    /// </summary>
    public AccessibleObject? GetFocusedObject()
    {
        // Spróbuj najpierw preferowanego API
        if (_providers.TryGetValue(_preferredApi, out var preferredProvider) && preferredProvider.IsActive)
        {
            var obj = preferredProvider.GetFocusedObject();
            if (obj != null)
                return obj;
        }

        // Spróbuj innych włączonych API
        foreach (var kvp in _providers)
        {
            if (kvp.Key != _preferredApi && _enabledApis.HasFlag(kvp.Key) && kvp.Value.IsActive)
            {
                var obj = kvp.Value.GetFocusedObject();
                if (obj != null)
                    return obj;
            }
        }

        return null;
    }

    /// <summary>
    /// Pobiera obiekt pod punktem ekranu
    /// </summary>
    public AccessibleObject? GetObjectFromPoint(int x, int y)
    {
        // Preferowane API pierwsze
        if (_providers.TryGetValue(_preferredApi, out var preferredProvider) && preferredProvider.IsActive)
        {
            var obj = preferredProvider.GetObjectFromPoint(x, y);
            if (obj != null)
                return obj;
        }

        // Inne API
        foreach (var kvp in _providers)
        {
            if (kvp.Key != _preferredApi && _enabledApis.HasFlag(kvp.Key) && kvp.Value.IsActive)
            {
                var obj = kvp.Value.GetObjectFromPoint(x, y);
                if (obj != null)
                    return obj;
            }
        }

        return null;
    }

    /// <summary>
    /// Pobiera najlepszy obiekt dla elementu UIA
    /// </summary>
    public AccessibleObject? GetAccessibleObject(AutomationElement element)
    {
        // Sprawdź każde API w kolejności preferencji
        if (_providers.TryGetValue(_preferredApi, out var preferredProvider) && preferredProvider.IsActive)
        {
            if (preferredProvider.SupportsElement(element))
            {
                var obj = preferredProvider.GetAccessibleObject(element);
                if (obj != null)
                    return obj;
            }
        }

        // Spróbuj innych API
        foreach (var kvp in _providers)
        {
            if (kvp.Key != _preferredApi && _enabledApis.HasFlag(kvp.Key) && kvp.Value.IsActive)
            {
                if (kvp.Value.SupportsElement(element))
                {
                    var obj = kvp.Value.GetAccessibleObject(element);
                    if (obj != null)
                        return obj;
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Pobiera provider dla danego API
    /// </summary>
    public IAccessibilityProvider? GetProvider(AccessibilityAPI api)
    {
        return _providers.TryGetValue(api, out var provider) ? provider : null;
    }

    /// <summary>
    /// Pobiera listę dostępnych API
    /// </summary>
    public IEnumerable<(AccessibilityAPI Api, bool IsAvailable, bool IsEnabled)> GetAvailableApis()
    {
        foreach (var kvp in _providers)
        {
            yield return (kvp.Key, kvp.Value.IsAvailable, _enabledApis.HasFlag(kvp.Key));
        }
    }

    /// <summary>
    /// Rozpoczyna nasłuchiwanie zdarzeń wszystkich włączonych providerów
    /// </summary>
    public void StartEventListening()
    {
        foreach (var kvp in _providers)
        {
            if (_enabledApis.HasFlag(kvp.Key) && kvp.Value.IsActive)
            {
                kvp.Value.StartEventListening();
            }
        }
    }

    /// <summary>
    /// Zatrzymuje nasłuchiwanie zdarzeń
    /// </summary>
    public void StopEventListening()
    {
        foreach (var provider in _providers.Values)
        {
            provider.StopEventListening();
        }
    }

    private void OnProviderFocusChanged(object? sender, AccessibleObject obj)
    {
        FocusChanged?.Invoke(this, obj);
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        StopEventListening();

        foreach (var provider in _providers.Values)
        {
            provider.Dispose();
        }

        _providers.Clear();
        _disposed = true;
    }
}
