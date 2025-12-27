using System.Diagnostics;
using System.Windows.Automation;

namespace ScreenReader.AppModules;

/// <summary>
/// Zarządza modułami aplikacji
/// Port z NVDA appModuleHandler.py
/// </summary>
public class AppModuleManager : IDisposable
{
    private readonly Dictionary<string, Func<AppModuleBase>> _moduleFactories = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<int, AppModuleBase> _activeModules = new();
    private AppModuleBase? _currentModule;
    private int _currentProcessId;

    public AppModuleBase? CurrentModule => _currentModule;

    public AppModuleManager()
    {
        RegisterBuiltInModules();
    }

    /// <summary>
    /// Rejestruje wbudowane moduły aplikacji
    /// </summary>
    private void RegisterBuiltInModules()
    {
        // Windows Explorer
        RegisterModule("explorer", () => new ExplorerModule());

        // Przeglądarki Chromium
        RegisterModule("chrome", () => new ChromeModule());
        RegisterModule("msedge", () => new MsEdgeModule());

        // Firefox (na przyszłość)
        // RegisterModule("firefox", () => new FirefoxModule());

        // UWP / Windows Store Apps
        RegisterModule("ApplicationFrameHost", () => new UWPModule("ApplicationFrameHost"));

        // Kalkulator Windows
        RegisterModule("Calculator", () => new CalculatorModule());
        RegisterModule("WindowsCalculator", () => new CalculatorModule());
        RegisterModule("Microsoft.WindowsCalculator", () => new CalculatorModule());
        RegisterModule("CalculatorApp", () => new CalculatorModule());

        // Ustawienia Windows
        RegisterModule("SystemSettings", () => new SettingsModule());
        RegisterModule("ms-settings", () => new SettingsModule());

        // Terminale - specjalna nawigacja NumPad
        RegisterModule("WindowsTerminal", () => new WindowsTerminalModule());
        RegisterModule("Microsoft.WindowsTerminal", () => new WindowsTerminalModule());
        RegisterModule("cmd", () => new CmdModule());
        RegisterModule("powershell", () => new PowerShellModule());
        RegisterModule("pwsh", () => new PwshModule());
        RegisterModule("conhost", () => new CmdModule()); // Console Host

        // Inne popularne aplikacje UWP
        RegisterModule("Photos", () => new UWPModule("Photos"));
        RegisterModule("Microsoft.Photos", () => new UWPModule("Photos"));
        RegisterModule("Microsoft.WindowsStore", () => new UWPModule("WindowsStore"));

        Console.WriteLine($"AppModuleManager: Zarejestrowano {_moduleFactories.Count} modułów");
    }

    /// <summary>
    /// Sprawdza czy aktualny moduł to terminal
    /// </summary>
    public bool IsTerminalActive => _currentModule is TerminalModule;

    /// <summary>
    /// Pobiera moduł terminala (jeśli aktywny)
    /// </summary>
    public TerminalModule? GetTerminalModule()
    {
        return _currentModule as TerminalModule;
    }

    /// <summary>
    /// Rejestruje fabrykę modułu dla danej nazwy procesu
    /// </summary>
    public void RegisterModule(string processName, Func<AppModuleBase> factory)
    {
        _moduleFactories[processName.ToLowerInvariant()] = factory;
    }

    /// <summary>
    /// Pobiera lub tworzy moduł dla danego procesu
    /// </summary>
    public AppModuleBase? GetModuleForProcess(int processId)
    {
        // Sprawdź czy mamy już moduł dla tego procesu
        if (_activeModules.TryGetValue(processId, out var existing))
            return existing;

        // Pobierz nazwę procesu
        string? processName = GetProcessName(processId);
        if (string.IsNullOrEmpty(processName))
            return null;

        // Normalizuj nazwę procesu
        processName = NormalizeProcessName(processName);

        // Sprawdź czy mamy fabrykę dla tego procesu
        if (!_moduleFactories.TryGetValue(processName, out var factory))
            return null;

        // Utwórz nowy moduł
        var module = factory();
        module.ProcessId = processId;
        _activeModules[processId] = module;

        Console.WriteLine($"AppModuleManager: Utworzono moduł {processName} dla PID {processId}");
        return module;
    }

    /// <summary>
    /// Obsługuje zmianę okna/fokusa
    /// </summary>
    public void OnFocusChanged(AutomationElement element)
    {
        int processId;
        try
        {
            processId = element.Current.ProcessId;
        }
        catch
        {
            return;
        }

        // Zmiana procesu
        if (processId != _currentProcessId)
        {
            // Powiadom stary moduł o utracie fokusu
            _currentModule?.OnLoseFocus();

            // Pobierz nowy moduł
            _currentModule = GetModuleForProcess(processId);
            _currentProcessId = processId;

            // Powiadom nowy moduł o uzyskaniu fokusu
            _currentModule?.OnGainFocus(element);
        }
        else
        {
            // Ten sam proces - powiadom o zmianie fokusu
            _currentModule?.OnFocusChanged(element);
        }
    }

    /// <summary>
    /// Obsługuje zmianę okna
    /// </summary>
    public void OnWindowChanged(AutomationElement window)
    {
        _currentModule?.OnWindowChanged(window);
    }

    /// <summary>
    /// Pobiera nazwę procesu z ID
    /// </summary>
    private static string? GetProcessName(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return Path.GetFileNameWithoutExtension(process.ProcessName);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Normalizuje nazwę procesu (obsługuje specjalne przypadki)
    /// </summary>
    private static string NormalizeProcessName(string processName)
    {
        // Obsługa różnych wersji aplikacji
        var normalized = processName.ToLowerInvariant();

        // Visual Studio Code variants
        if (normalized.StartsWith("code"))
            return "code";

        // Microsoft Edge
        if (normalized.Contains("msedge"))
            return "msedge";

        return normalized;
    }

    /// <summary>
    /// Czyści nieaktywne moduły
    /// </summary>
    public void CleanupInactiveModules()
    {
        var toRemove = new List<int>();

        foreach (var kvp in _activeModules)
        {
            try
            {
                using var process = Process.GetProcessById(kvp.Key);
                if (process.HasExited)
                    toRemove.Add(kvp.Key);
            }
            catch
            {
                toRemove.Add(kvp.Key);
            }
        }

        foreach (var pid in toRemove)
        {
            if (_activeModules.TryGetValue(pid, out var module))
            {
                module.Dispose();
                _activeModules.Remove(pid);
            }
        }
    }

    public void Dispose()
    {
        foreach (var module in _activeModules.Values)
        {
            module.Dispose();
        }
        _activeModules.Clear();
    }
}
