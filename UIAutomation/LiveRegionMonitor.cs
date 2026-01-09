using System.Runtime.InteropServices;
using System.Windows.Automation;

namespace ScreenReader;

/// <summary>
/// Monitoruje LiveRegion i NotificationEvent dla komunikatów dostępności
/// Obsługuje MSAA, UIA i inne metody dostępności
/// Rozbudowane monitorowanie dynamicznych treści
/// </summary>
public class LiveRegionMonitor : IDisposable
{
    private readonly object _lock = new();
    private AutomationEventHandler? _notificationHandler;
    private IntPtr _winEventHook;
    private IntPtr _textChangeHook;
    private IntPtr _nameChangeHook;
    private IntPtr _valueChangeHook;
    private readonly WinEventDelegate _winEventDelegate;
    private int _targetProcessId;
    private IntPtr _targetWindowHandle;
    private bool _disposed;
    private bool _isRunning;
    private bool _monitorAllChanges = true;
    private bool _onlyActiveWindow = false; // Wyłączone domyślnie - użytkownik może włączyć

    // Cache dla wykrywania powtórzeń
    private readonly Dictionary<string, DateTime> _recentAnnouncements = new();
    private readonly TimeSpan _deduplicationWindow = TimeSpan.FromMilliseconds(500);

    // Event dla komunikatów LiveRegion
    public event Action<string, bool>? LiveRegionChanged;

    // Event dla zmiany tekstu (ogólne)
    public event Action<string, TextChangeType>? TextChanged;

    // Event dla zmiany struktury (nowe elementy)
    public event Action<string>? StructureChanged;

    /// <summary>
    /// Czy monitorować wszystkie zmiany tekstu (nie tylko LiveRegion)
    /// </summary>
    public bool MonitorAllChanges
    {
        get => _monitorAllChanges;
        set
        {
            _monitorAllChanges = value;
            Console.WriteLine($"LiveRegionMonitor: Monitorowanie wszystkich zmian: {value}");
        }
    }

    /// <summary>
    /// Czy monitorować tylko aktywne okno (domyślnie true)
    /// </summary>
    public bool OnlyActiveWindow
    {
        get => _onlyActiveWindow;
        set
        {
            _onlyActiveWindow = value;
            Console.WriteLine($"LiveRegionMonitor: Monitorowanie tylko aktywnego okna: {value}");
        }
    }

    // WinAPI Constants
    private const uint EVENT_OBJECT_LIVEREGIONCHANGED = 0x8019;
    private const uint EVENT_OBJECT_NAMECHANGE = 0x800C;
    private const uint EVENT_OBJECT_VALUECHANGE = 0x800E;
    private const uint EVENT_OBJECT_TEXTSELECTIONCHANGED = 0x8014;
    private const uint EVENT_OBJECT_CONTENTSCROLLED = 0x8015;
    private const uint EVENT_OBJECT_TEXTEDIT_CONVERSIONTARGETCHANGED = 0x8030;
    private const uint EVENT_SYSTEM_ALERT = 0x0002;
    private const uint WINEVENT_OUTOFCONTEXT = 0x0000;
    private const int OBJID_CLIENT = -4;

    // UIA NotificationEvent - ID 20036
    private static readonly AutomationEvent? NotificationEvent;
    private static readonly AutomationProperty? LiveSettingProperty;

    private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

    [DllImport("user32.dll")]
    private static extern IntPtr SetWinEventHook(
        uint eventMin, uint eventMax, IntPtr hmodWinEventProc,
        WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

    [DllImport("user32.dll")]
    private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(
        IntPtr hwnd, uint dwId, ref Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out object ppvObject);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    static LiveRegionMonitor()
    {
        try
        {
            // UIA_NotificationEventId = 20036
            NotificationEvent = AutomationEvent.LookupById(20036);
            // UIA_LiveSettingPropertyId = 30135
            LiveSettingProperty = AutomationProperty.LookupById(30135);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"LiveRegionMonitor: Nie można zainicjować UIA events: {ex.Message}");
        }
    }

    public LiveRegionMonitor()
    {
        _winEventDelegate = OnWinEvent;
    }

    /// <summary>
    /// Rozpoczyna monitorowanie LiveRegion
    /// </summary>
    public void Start()
    {
        if (_isRunning)
            return;

        _isRunning = true;

        // Hook Win32 events dla MSAA LiveRegion
        _winEventHook = SetWinEventHook(
            EVENT_OBJECT_LIVEREGIONCHANGED, EVENT_OBJECT_LIVEREGIONCHANGED,
            IntPtr.Zero, _winEventDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);

        if (_winEventHook != IntPtr.Zero)
        {
            Console.WriteLine("LiveRegionMonitor: MSAA LiveRegion hook zainstalowany");
        }

        // Hook dla zmian nazwy (NAME_CHANGE)
        _nameChangeHook = SetWinEventHook(
            EVENT_OBJECT_NAMECHANGE, EVENT_OBJECT_NAMECHANGE,
            IntPtr.Zero, _winEventDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);

        if (_nameChangeHook != IntPtr.Zero)
        {
            Console.WriteLine("LiveRegionMonitor: NameChange hook zainstalowany");
        }

        // Hook dla zmian wartości (VALUE_CHANGE)
        _valueChangeHook = SetWinEventHook(
            EVENT_OBJECT_VALUECHANGE, EVENT_OBJECT_VALUECHANGE,
            IntPtr.Zero, _winEventDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);

        if (_valueChangeHook != IntPtr.Zero)
        {
            Console.WriteLine("LiveRegionMonitor: ValueChange hook zainstalowany");
        }

        // Hook dla alertów systemowych
        _textChangeHook = SetWinEventHook(
            EVENT_SYSTEM_ALERT, EVENT_SYSTEM_ALERT,
            IntPtr.Zero, _winEventDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);

        // Subskrybuj UIA NotificationEvent jeśli dostępny
        if (NotificationEvent != null)
        {
            try
            {
                _notificationHandler = OnUiaNotification;
                Automation.AddAutomationEventHandler(
                    NotificationEvent,
                    AutomationElement.RootElement,
                    TreeScope.Subtree,
                    _notificationHandler);

                Console.WriteLine("LiveRegionMonitor: UIA NotificationEvent handler dodany");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"LiveRegionMonitor: Błąd UIA Notification: {ex.Message}");
            }
        }

        // Subskrybuj UIA TextChanged event
        try
        {
            Automation.AddAutomationEventHandler(
                TextPattern.TextChangedEvent,
                AutomationElement.RootElement,
                TreeScope.Subtree,
                OnUiaTextChanged);

            Console.WriteLine("LiveRegionMonitor: UIA TextChanged handler dodany");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"LiveRegionMonitor: Błąd UIA TextChanged: {ex.Message}");
        }

        // Subskrybuj UIA StructureChanged event
        try
        {
            Automation.AddStructureChangedEventHandler(
                AutomationElement.RootElement,
                TreeScope.Subtree,
                OnUiaStructureChanged);

            Console.WriteLine("LiveRegionMonitor: UIA StructureChanged handler dodany");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"LiveRegionMonitor: Błąd UIA StructureChanged: {ex.Message}");
        }

        Console.WriteLine("LiveRegionMonitor: Monitorowanie rozpoczęte");
    }

    /// <summary>
    /// Obsługa UIA TextChanged event
    /// </summary>
    private void OnUiaTextChanged(object sender, AutomationEventArgs e)
    {
        if (!_monitorAllChanges)
            return;

        try
        {
            if (sender is not AutomationElement element)
                return;

            // Sprawdź czy event jest z aktywnego okna
            if (!IsFromActiveWindow(element))
                return;

            // Filtruj po procesie jeśli ustawiony
            int targetPid;
            lock (_lock)
            {
                targetPid = _targetProcessId;
            }

            if (targetPid > 0)
            {
                try
                {
                    if (element.Current.ProcessId != targetPid)
                        return;
                }
                catch { return; }
            }

            string text = GetElementText(element);
            if (!string.IsNullOrWhiteSpace(text) && !IsDuplicate(text))
            {
                Console.WriteLine($"LiveRegionMonitor (TextChanged): {text}");
                TextChanged?.Invoke(text, TextChangeType.ContentChanged);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"LiveRegionMonitor: Błąd UIA TextChanged: {ex.Message}");
        }
    }

    /// <summary>
    /// Obsługa UIA StructureChanged event
    /// </summary>
    private void OnUiaStructureChanged(object sender, StructureChangedEventArgs e)
    {
        if (!_monitorAllChanges)
            return;

        try
        {
            if (sender is not AutomationElement element)
                return;

            // Tylko dla dodanych elementów
            if (e.StructureChangeType != StructureChangeType.ChildAdded &&
                e.StructureChangeType != StructureChangeType.ChildrenBulkAdded)
                return;

            // Sprawdź czy event jest z aktywnego okna
            if (!IsFromActiveWindow(element))
                return;

            // Filtruj po procesie jeśli ustawiony
            int targetPid;
            lock (_lock)
            {
                targetPid = _targetProcessId;
            }

            if (targetPid > 0)
            {
                try
                {
                    if (element.Current.ProcessId != targetPid)
                        return;
                }
                catch { return; }
            }

            // Sprawdź czy to live region
            var liveSettingValue = GetLiveSetting(element);
            if (liveSettingValue > 0)
            {
                string text = GetElementText(element);
                if (!string.IsNullOrWhiteSpace(text) && !IsDuplicate(text))
                {
                    Console.WriteLine($"LiveRegionMonitor (StructureChanged): {text}");
                    StructureChanged?.Invoke(text);
                    LiveRegionChanged?.Invoke(text, liveSettingValue == 2); // 2 = assertive
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"LiveRegionMonitor: Błąd UIA StructureChanged: {ex.Message}");
        }
    }

    /// <summary>
    /// Pobiera wartość LiveSetting z elementu
    /// </summary>
    private static int GetLiveSetting(AutomationElement element)
    {
        try
        {
            if (LiveSettingProperty != null)
            {
                var value = element.GetCurrentPropertyValue(LiveSettingProperty);
                if (value is int intValue)
                    return intValue;
            }
        }
        catch { }

        return 0;
    }

    /// <summary>
    /// Sprawdza czy komunikat był niedawno ogłoszony (deduplikacja)
    /// </summary>
    private bool IsDuplicate(string text)
    {
        lock (_lock)
        {
            var now = DateTime.Now;

            // Wyczyść stare wpisy
            var keysToRemove = _recentAnnouncements
                .Where(kvp => now - kvp.Value > _deduplicationWindow)
                .Select(kvp => kvp.Key)
                .ToList();

            foreach (var key in keysToRemove)
            {
                _recentAnnouncements.Remove(key);
            }

            // Sprawdź czy tekst był niedawno ogłoszony
            if (_recentAnnouncements.ContainsKey(text))
                return true;

            // Dodaj do cache
            _recentAnnouncements[text] = now;
            return false;
        }
    }

    /// <summary>
    /// Zatrzymuje monitorowanie
    /// </summary>
    public void Stop()
    {
        if (!_isRunning)
            return;

        _isRunning = false;

        // Usuń Win32 hooki
        if (_winEventHook != IntPtr.Zero)
        {
            UnhookWinEvent(_winEventHook);
            _winEventHook = IntPtr.Zero;
        }

        if (_nameChangeHook != IntPtr.Zero)
        {
            UnhookWinEvent(_nameChangeHook);
            _nameChangeHook = IntPtr.Zero;
        }

        if (_valueChangeHook != IntPtr.Zero)
        {
            UnhookWinEvent(_valueChangeHook);
            _valueChangeHook = IntPtr.Zero;
        }

        if (_textChangeHook != IntPtr.Zero)
        {
            UnhookWinEvent(_textChangeHook);
            _textChangeHook = IntPtr.Zero;
        }

        // Usuń UIA handlers
        try
        {
            if (_notificationHandler != null && NotificationEvent != null)
            {
                Automation.RemoveAutomationEventHandler(
                    NotificationEvent,
                    AutomationElement.RootElement,
                    _notificationHandler);
            }

            Automation.RemoveAllEventHandlers();
        }
        catch { }

        Console.WriteLine("LiveRegionMonitor: Monitorowanie zatrzymane");
    }

    /// <summary>
    /// Ustawia docelowy proces (dla filtrowania eventów)
    /// </summary>
    public void SetTargetProcess(int processId)
    {
        lock (_lock)
        {
            _targetProcessId = processId;
        }
    }

    /// <summary>
    /// Ustawia docelowe okno dla filtrowania eventów (monitoruje tylko aktywne okno)
    /// </summary>
    public void SetTargetWindow(IntPtr hwnd)
    {
        lock (_lock)
        {
            _targetWindowHandle = hwnd;
            if (hwnd != IntPtr.Zero)
            {
                GetWindowThreadProcessId(hwnd, out uint pid);
                _targetProcessId = (int)pid;
            }
        }
    }

    /// <summary>
    /// Aktualizuje target na bieżące okno pierwszego planu
    /// </summary>
    public void UpdateToForegroundWindow()
    {
        if (!_onlyActiveWindow)
            return;

        var hwnd = GetForegroundWindow();
        if (hwnd != IntPtr.Zero)
        {
            SetTargetWindow(hwnd);
        }
    }

    /// <summary>
    /// Sprawdza czy element należy do aktywnego okna
    /// </summary>
    private bool IsFromActiveWindow(AutomationElement element)
    {
        if (!_onlyActiveWindow)
            return true;

        try
        {
            IntPtr foregroundHwnd = GetForegroundWindow();
            if (foregroundHwnd == IntPtr.Zero)
                return true;

            GetWindowThreadProcessId(foregroundHwnd, out uint foregroundPid);

            // Sprawdź czy element jest z tego samego procesu
            int elementPid = element.Current.ProcessId;
            return elementPid == (int)foregroundPid;
        }
        catch
        {
            return true; // W razie błędu pozwól na przejście
        }
    }

    /// <summary>
    /// Sprawdza czy hwnd należy do aktywnego okna/procesu
    /// </summary>
    private bool IsFromActiveWindow(IntPtr hwnd)
    {
        if (!_onlyActiveWindow)
            return true;

        if (hwnd == IntPtr.Zero)
            return false;

        try
        {
            IntPtr foregroundHwnd = GetForegroundWindow();
            if (foregroundHwnd == IntPtr.Zero)
                return true;

            // Sprawdź czy to to samo okno
            if (hwnd == foregroundHwnd)
                return true;

            // Sprawdź czy to ten sam proces
            GetWindowThreadProcessId(foregroundHwnd, out uint foregroundPid);
            GetWindowThreadProcessId(hwnd, out uint eventPid);

            return eventPid == foregroundPid;
        }
        catch
        {
            return true;
        }
    }

    /// <summary>
    /// Obsługa Win32 WinEvent (MSAA LiveRegion, NameChange, ValueChange)
    /// </summary>
    private void OnWinEvent(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        try
        {
            // Najpierw sprawdź czy event jest z aktywnego okna (szybkie odrzucenie)
            if (!IsFromActiveWindow(hwnd))
                return;

            // Użyj UI Automation żeby pobrać element z hwnd
            var element = AutomationElement.FromHandle(hwnd);
            if (element == null)
                return;

            // Sprawdź czy to z docelowego procesu (dodatkowa weryfikacja)
            int targetPid;
            lock (_lock)
            {
                targetPid = _targetProcessId;
            }

            if (targetPid > 0)
            {
                try
                {
                    if (element.Current.ProcessId != targetPid)
                        return;
                }
                catch { return; }
            }

            // Pobierz tekst
            string text = GetElementText(element);
            if (string.IsNullOrWhiteSpace(text))
                return;

            // Sprawdź deduplikację
            if (IsDuplicate(text))
                return;

            switch (eventType)
            {
                case EVENT_OBJECT_LIVEREGIONCHANGED:
                    Console.WriteLine($"LiveRegionMonitor (MSAA LiveRegion): {text}");
                    LiveRegionChanged?.Invoke(text, true);
                    break;

                case EVENT_OBJECT_NAMECHANGE:
                    if (_monitorAllChanges)
                    {
                        // Sprawdź czy to live region
                        var liveSetting = GetLiveSetting(element);
                        if (liveSetting > 0)
                        {
                            Console.WriteLine($"LiveRegionMonitor (NameChange/LiveRegion): {text}");
                            LiveRegionChanged?.Invoke(text, liveSetting == 2);
                        }
                        else
                        {
                            Console.WriteLine($"LiveRegionMonitor (NameChange): {text}");
                            TextChanged?.Invoke(text, TextChangeType.NameChanged);
                        }
                    }
                    break;

                case EVENT_OBJECT_VALUECHANGE:
                    if (_monitorAllChanges)
                    {
                        var liveSetting = GetLiveSetting(element);
                        if (liveSetting > 0)
                        {
                            Console.WriteLine($"LiveRegionMonitor (ValueChange/LiveRegion): {text}");
                            LiveRegionChanged?.Invoke(text, liveSetting == 2);
                        }
                        else
                        {
                            Console.WriteLine($"LiveRegionMonitor (ValueChange): {text}");
                            TextChanged?.Invoke(text, TextChangeType.ValueChanged);
                        }
                    }
                    break;

                case EVENT_SYSTEM_ALERT:
                    Console.WriteLine($"LiveRegionMonitor (SystemAlert): {text}");
                    LiveRegionChanged?.Invoke(text, true); // Alerty są zawsze assertive
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"LiveRegionMonitor: Błąd MSAA event {eventType}: {ex.Message}");
        }
    }

    /// <summary>
    /// Obsługa UIA NotificationEvent
    /// </summary>
    private void OnUiaNotification(object sender, AutomationEventArgs e)
    {
        try
        {
            if (sender is not AutomationElement element)
                return;

            // Sprawdź czy event jest z aktywnego okna
            if (!IsFromActiveWindow(element))
                return;

            // Pobierz tekst powiadomienia
            string text = GetElementText(element);
            if (string.IsNullOrWhiteSpace(text))
                return;

            Console.WriteLine($"LiveRegionMonitor (Notification): {text}");
            LiveRegionChanged?.Invoke(text, true);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"LiveRegionMonitor: Błąd UIA Notification: {ex.Message}");
        }
    }

    /// <summary>
    /// Pobiera tekst z elementu
    /// </summary>
    private static string GetElementText(AutomationElement element)
    {
        try
        {
            // Próbuj pobrać nazwę
            var name = element.Current.Name;
            if (!string.IsNullOrWhiteSpace(name))
                return name;

            // Próbuj ValuePattern
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var valuePattern))
            {
                var value = ((ValuePattern)valuePattern).Current.Value;
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            // Próbuj TextPattern
            if (element.TryGetCurrentPattern(TextPattern.Pattern, out var textPattern))
            {
                var text = ((TextPattern)textPattern).DocumentRange.GetText(-1);
                if (!string.IsNullOrWhiteSpace(text))
                    return text;
            }
        }
        catch { }

        return "";
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        Stop();
        _disposed = true;
    }
}

/// <summary>
/// Typ zmiany tekstu
/// </summary>
public enum TextChangeType
{
    /// <summary>
    /// Zmiana zawartości (treść)
    /// </summary>
    ContentChanged,

    /// <summary>
    /// Zmiana nazwy elementu
    /// </summary>
    NameChanged,

    /// <summary>
    /// Zmiana wartości elementu
    /// </summary>
    ValueChanged,

    /// <summary>
    /// Nowy element dodany
    /// </summary>
    ElementAdded,

    /// <summary>
    /// Alert systemowy
    /// </summary>
    SystemAlert
}
