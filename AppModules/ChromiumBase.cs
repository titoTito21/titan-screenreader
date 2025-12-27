using System.Windows.Automation;
using ScreenReader.Interop;

namespace ScreenReader.AppModules;

/// <summary>
/// Bazowy moduł dla przeglądarek Chromium (Chrome, Edge)
/// Port z NVDA appModules - wspólny backend dla Chromium
///
/// Obsługuje automatyczne przełączanie między modelami dostępności:
/// - UIA (UI Automation) - domyślny w Chrome 138+
/// - IAccessible2 - legacy, ale bardziej kompletny dla wirtualnego bufora
/// </summary>
public class ChromiumBaseModule : AppModuleBase
{
    /// <summary>Manager modeli dostępności</summary>
    protected readonly AccessibilityModelManager AccessibilityManager;

    /// <summary>Czy dostępność została aktywowana dla tego procesu</summary>
    protected bool AccessibilityActivated { get; private set; }

    /// <summary>Aktualny model dostępności</summary>
    public AccessibilityModel CurrentAccessibilityModel =>
        AccessibilityManager.GetModelForProcess(ProcessId);

    protected ChromiumBaseModule(string processName) : base(processName)
    {
        AccessibilityManager = new AccessibilityModelManager();
        AccessibilityManager.ModelChanged += OnAccessibilityModelChanged;
    }

    public override void OnGainFocus(AutomationElement element)
    {
        base.OnGainFocus(element);

        // Automatycznie aktywuj dostępność IAccessible2 dla przeglądarki
        if (!AccessibilityActivated && AccessibilityManager.AutoActivateIA2)
        {
            ActivateAccessibility();
        }

        if (ShouldUseVirtualBuffer(element))
        {
            Console.WriteLine($"{ProcessName}: Aktywacja trybu przeglądania (model: {CurrentAccessibilityModel})");
        }
    }

    /// <summary>
    /// Aktywuje interfejsy dostępności dla przeglądarki
    /// </summary>
    protected virtual void ActivateAccessibility()
    {
        if (ProcessId <= 0)
            return;

        Console.WriteLine($"{ProcessName}: Aktywacja dostępności...");

        bool success = AccessibilityManager.ActivateAccessibility(ProcessId);
        AccessibilityActivated = success;

        if (success)
        {
            Console.WriteLine($"{ProcessName}: Dostępność aktywowana pomyślnie");
        }
        else
        {
            Console.WriteLine($"{ProcessName}: Nie udało się aktywować IAccessible2, używam UIA");
        }
    }

    /// <summary>
    /// Obsługa zmiany modelu dostępności
    /// </summary>
    protected virtual void OnAccessibilityModelChanged(int processId, AccessibilityModel model)
    {
        if (processId == ProcessId)
        {
            Console.WriteLine($"{ProcessName}: Model dostępności zmieniony na {model}");
        }
    }

    public override void OnFocusChanged(AutomationElement element)
    {
        var controlType = element.Current.ControlType;

        // Wykryj zmianę zakładki
        if (controlType == ControlType.TabItem)
        {
            HandleTabSwitch(element);
        }
        // Wykryj pasek adresu
        else if (IsAddressBar(element))
        {
            Console.WriteLine($"{ProcessName}: Pasek adresu");
        }
    }

    /// <summary>
    /// Czy używać wirtualnego bufora dla tego elementu
    /// </summary>
    public override bool ShouldUseVirtualBuffer(AutomationElement element)
    {
        // Użyj wirtualnego bufora dla dokumentów webowych
        return element.Current.ControlType == ControlType.Document;
    }

    /// <summary>
    /// Obsługuje zmianę zakładki
    /// </summary>
    protected virtual void HandleTabSwitch(AutomationElement tab)
    {
        try
        {
            string title = tab.Current.Name;
            string? url = GetAddressBarText();
            string? domain = GetDomain(url);

            string announcement = title;
            if (!string.IsNullOrEmpty(domain))
                announcement += $", {domain}";

            Console.WriteLine($"{ProcessName}: Zakładka {announcement}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"{ProcessName}: Błąd przy zmianie zakładki: {ex.Message}");
        }
    }

    /// <summary>
    /// Pobiera tekst paska adresu
    /// </summary>
    public string? GetAddressBarText()
    {
        try
        {
            var root = AutomationElement.RootElement;
            var chromeWindow = FindChromeWindow(root);

            if (chromeWindow != null)
            {
                var addressBar = FindAddressBar(chromeWindow);
                if (addressBar != null)
                {
                    if (addressBar.TryGetCurrentPattern(ValuePattern.Pattern, out var pattern))
                    {
                        return ((ValuePattern)pattern).Current.Value;
                    }
                    return addressBar.Current.Name;
                }
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Znajduje okno przeglądarki
    /// </summary>
    protected AutomationElement? FindChromeWindow(AutomationElement root)
    {
        try
        {
            var condition = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                new PropertyCondition(AutomationElement.ProcessIdProperty, ProcessId)
            );

            return root.FindFirst(TreeScope.Children, condition);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Znajduje pasek adresu w oknie przeglądarki
    /// </summary>
    protected AutomationElement? FindAddressBar(AutomationElement window)
    {
        try
        {
            // Chromium używa Edit control dla paska adresu
            var condition = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                new PropertyCondition(AutomationElement.AutomationIdProperty, "addressBarEdit")
            );

            var addressBar = window.FindFirst(TreeScope.Descendants, condition);
            if (addressBar != null)
                return addressBar;

            // Fallback: szukaj po nazwie
            condition = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                new PropertyCondition(AutomationElement.NameProperty, "Adres i pasek wyszukiwania")
            );

            return window.FindFirst(TreeScope.Descendants, condition);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Sprawdza czy element to pasek adresu
    /// </summary>
    protected bool IsAddressBar(AutomationElement element)
    {
        try
        {
            if (element.Current.ControlType != ControlType.Edit)
                return false;

            string name = element.Current.Name;
            string automationId = element.Current.AutomationId;

            return automationId == "addressBarEdit" ||
                   name.Contains("adres", StringComparison.OrdinalIgnoreCase) ||
                   name.Contains("address", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Wyciąga domenę z URL
    /// </summary>
    protected string? GetDomain(string? url)
    {
        if (string.IsNullOrEmpty(url))
            return null;

        try
        {
            if (Uri.TryCreate(url, UriKind.Absolute, out var uri))
            {
                return uri.Host;
            }

            // Jeśli URL nie ma protokołu, dodaj https
            if (!url.Contains("://"))
            {
                if (Uri.TryCreate("https://" + url, UriKind.Absolute, out uri))
                {
                    return uri.Host;
                }
            }
        }
        catch { }

        return null;
    }

    /// <summary>
    /// Obsługuje powiadomienia przeglądarki
    /// </summary>
    public override void HandleNotification(string text)
    {
        Console.WriteLine($"{ProcessName}: Powiadomienie - {text}");
    }

    public override string? GetStatusBarText()
    {
        // Chromium nie ma tradycyjnego paska stanu
        // Można zwrócić URL z paska adresu
        return GetAddressBarText();
    }
}

/// <summary>
/// Moduł dla Google Chrome
/// </summary>
public class ChromeModule : ChromiumBaseModule
{
    public ChromeModule() : base("chrome")
    {
    }
}

/// <summary>
/// Moduł dla Microsoft Edge
/// Port z NVDA appModules/msedge.py
///
/// Edge na Windows 11 wymaga specjalnej obsługi:
/// - Automatyczna aktywacja IAccessible2 dla lepszej obsługi wirtualnego bufora
/// - Ulepszone wykrywanie dokumentów webowych
/// - Obsługa specyficznych elementów Edge (kolekcje, pasek boczny)
/// </summary>
public class MsEdgeModule : ChromiumBaseModule
{
    /// <summary>Czy to jest Windows 11</summary>
    private readonly bool _isWindows11;

    /// <summary>Ostatni czas odświeżenia bufora</summary>
    private DateTime _lastBufferRefresh = DateTime.MinValue;

    /// <summary>Minimalny czas między odświeżeniami bufora (ms)</summary>
    private const int MinBufferRefreshInterval = 500;

    public MsEdgeModule() : base("msedge")
    {
        // Sprawdź wersję Windows
        _isWindows11 = Environment.OSVersion.Version.Build >= 22000;

        if (_isWindows11)
        {
            Console.WriteLine("MsEdge: Wykryto Windows 11, używam ulepszonej obsługi");
        }
    }

    public override void OnGainFocus(AutomationElement element)
    {
        // Na Windows 11 aktywuj dostępność wcześniej
        if (_isWindows11 && !AccessibilityActivated)
        {
            ActivateAccessibility();
        }

        base.OnGainFocus(element);
    }

    protected override void ActivateAccessibility()
    {
        base.ActivateAccessibility();

        // Na Windows 11 spróbuj też aktywować okno z treścią
        if (_isWindows11 && ProcessId > 0)
        {
            try
            {
                using var process = System.Diagnostics.Process.GetProcessById(ProcessId);
                var mainWindow = process.MainWindowHandle;
                if (mainWindow != IntPtr.Zero)
                {
                    AccessibilityManager.ActivateContentWindow(mainWindow);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"MsEdge: Błąd aktywacji okna treści: {ex.Message}");
            }
        }
    }

    public override bool ShouldUseVirtualBuffer(AutomationElement element)
    {
        try
        {
            var controlType = element.Current.ControlType;

            // Standardowe dokumenty
            if (controlType == ControlType.Document)
                return true;

            // Edge czasem używa Pane dla treści strony
            if (controlType == ControlType.Pane)
            {
                var className = element.Current.ClassName;
                // Chrome_RenderWidgetHostHWND to klasa okna z treścią strony
                if (!string.IsNullOrEmpty(className) &&
                    (className.Contains("Chrome") || className.Contains("RenderWidget")))
                {
                    return true;
                }

                // Sprawdź czy ma dzieci typu Document
                var walker = TreeWalker.ControlViewWalker;
                var child = walker.GetFirstChild(element);
                while (child != null)
                {
                    if (child.Current.ControlType == ControlType.Document)
                        return true;
                    child = walker.GetNextSibling(child);
                }
            }

            return false;
        }
        catch
        {
            return false;
        }
    }

    protected override void HandleTabSwitch(AutomationElement tab)
    {
        base.HandleTabSwitch(tab);

        // Na Windows 11 odśwież wirtualny bufor przy zmianie zakładki
        if (_isWindows11)
        {
            var now = DateTime.Now;
            if ((now - _lastBufferRefresh).TotalMilliseconds > MinBufferRefreshInterval)
            {
                _lastBufferRefresh = now;
                Console.WriteLine("MsEdge: Odświeżenie bufora po zmianie zakładki");
                // Tutaj można wywołać odświeżenie wirtualnego bufora
            }
        }
    }

    /// <summary>
    /// Sprawdza czy element to pasek boczny Edge
    /// </summary>
    public bool IsSidebarElement(AutomationElement element)
    {
        try
        {
            var name = element.Current.Name.ToLowerInvariant();
            return name.Contains("sidebar") ||
                   name.Contains("pasek boczny") ||
                   name.Contains("panel boczny");
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Sprawdza czy element to kolekcja Edge
    /// </summary>
    public bool IsCollectionElement(AutomationElement element)
    {
        try
        {
            var name = element.Current.Name.ToLowerInvariant();
            return name.Contains("collection") ||
                   name.Contains("kolekcj");
        }
        catch
        {
            return false;
        }
    }

    protected override void OnAccessibilityModelChanged(int processId, AccessibilityModel model)
    {
        base.OnAccessibilityModelChanged(processId, model);

        if (processId == ProcessId && model == AccessibilityModel.IAccessible2)
        {
            Console.WriteLine("MsEdge: IAccessible2 aktywne - wirtualny bufor będzie działał lepiej");
        }
    }
}
