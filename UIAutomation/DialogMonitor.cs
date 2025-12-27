using System.Runtime.InteropServices;
using System.Windows.Automation;

namespace ScreenReader;

public class DialogMonitor : IDisposable
{
    private IntPtr _hookHandle;
    private readonly SpeechManager _speechManager;
    private IntPtr _lastWindow = IntPtr.Zero;
    private readonly WinEventDelegate _eventDelegate;
    private bool _disposed;

    // WinAPI Constants
    private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
    private const uint EVENT_OBJECT_FOCUS = 0x8005;
    private const uint WINEVENT_OUTOFCONTEXT = 0x0000;
    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_DLGMODALFRAME = 0x00000001;
    private const int WS_EX_TOPMOST = 0x00000008;

    // Delegate for WinEvent callback
    private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

    // WinAPI Imports
    [DllImport("user32.dll")]
    private static extern IntPtr SetWinEventHook(
        uint eventMin, uint eventMax, IntPtr hmodWinEventProc,
        WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

    [DllImport("user32.dll")]
    private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    // Klasy okien do ignorowania
    private static readonly string[] IgnoredWindowClasses =
    {
        "tooltips_class",
        "SysShadow",
        "IME",
        "MSCTFIME UI",
        "Windows.UI.Core.CoreWindow", // Czasem są to ukryte okna
        "Shell_TrayWnd",
        "Shell_SecondaryTrayWnd",
        "NotifyIconOverflowWindow",
        "TaskListThumbnailWnd",
        "ForegroundStaging"
    };

    public DialogMonitor(SpeechManager speechManager)
    {
        _speechManager = speechManager;
        _eventDelegate = ProcessWindowEvent;
    }

    public void StartMonitoring()
    {
        Console.WriteLine("DialogMonitor: Rozpoczynam monitorowanie okien");
        
        // Hook window foreground events
        _hookHandle = SetWinEventHook(
            EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
            IntPtr.Zero, _eventDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);

        if (_hookHandle == IntPtr.Zero)
        {
            Console.WriteLine("DialogMonitor: Błąd instalacji hooka");
        }
        else
        {
            Console.WriteLine("DialogMonitor: Hook zainstalowany pomyślnie");
        }
    }

    private void ProcessWindowEvent(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        try
        {
            // Ignoruj duplikaty
            if (hwnd == _lastWindow || hwnd == IntPtr.Zero)
                return;

            // Sprawdź czy okno jest rzeczywiście na pierwszym planie (ma fokus)
            IntPtr foregroundWindow = GetForegroundWindow();
            if (hwnd != foregroundWindow)
            {
                Console.WriteLine("DialogMonitor: Okno nie ma fokusu, ignoruję");
                return;
            }

            // Sprawdź czy okno jest widoczne i nie zminimalizowane
            if (!IsWindowVisible(hwnd) || IsIconic(hwnd))
            {
                Console.WriteLine("DialogMonitor: Okno niewidoczne lub zminimalizowane, ignoruję");
                return;
            }

            // Pobierz klasę okna
            string className = GetWindowClassName(hwnd);

            // Ignoruj określone klasy okien
            if (IsIgnoredWindowClass(className))
            {
                Console.WriteLine($"DialogMonitor: Ignoruję klasę okna: {className}");
                return;
            }

            _lastWindow = hwnd;

            // Pobierz informacje o oknie
            string title = GetWindowTitle(hwnd);
            bool isDialog = IsDialog(hwnd);
            bool isTopmost = IsTopmost(hwnd);

            Console.WriteLine($"DialogMonitor: Nowe okno - Tytuł: {title}, Klasa: {className}, Dialog: {isDialog}, Topmost: {isTopmost}");

            // Jeśli to dialog, ogłoś specjalnie
            if (isDialog || isTopmost)
            {
                AnnounceDialog(hwnd, title, isTopmost);
            }
            else
            {
                AnnounceWindow(hwnd, title);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"DialogMonitor: Błąd przetwarzania wydarzenia: {ex.Message}");
        }
    }

    /// <summary>
    /// Sprawdza czy klasa okna powinna być ignorowana
    /// </summary>
    private bool IsIgnoredWindowClass(string className)
    {
        if (string.IsNullOrEmpty(className))
            return false;

        foreach (var ignoredClass in IgnoredWindowClasses)
        {
            if (className.Contains(ignoredClass, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    private bool IsDialog(IntPtr hwnd)
    {
        try
        {
            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            return (exStyle & WS_EX_DLGMODALFRAME) != 0;
        }
        catch
        {
            return false;
        }
    }

    private bool IsTopmost(IntPtr hwnd)
    {
        try
        {
            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            return (exStyle & WS_EX_TOPMOST) != 0;
        }
        catch
        {
            return false;
        }
    }

    private string GetWindowTitle(IntPtr hwnd)
    {
        try
        {
            int length = GetWindowTextLength(hwnd);
            if (length == 0)
                return "";

            var sb = new System.Text.StringBuilder(length + 1);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }
        catch
        {
            return "";
        }
    }

    private string GetWindowClassName(IntPtr hwnd)
    {
        try
        {
            var sb = new System.Text.StringBuilder(256);
            GetClassName(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }
        catch
        {
            return "";
        }
    }

    private void AnnounceDialog(IntPtr hwnd, string title, bool isTopmost)
    {
        string announcement = isTopmost ? "Okno dialogowe" : "Dialog";
        
        if (!string.IsNullOrEmpty(title))
        {
            _speechManager.Speak($"{announcement}: {title}", interrupt: true);
        }
        else
        {
            _speechManager.Speak(announcement, interrupt: true);
        }

        // Odczekaj chwilę i przeczytaj skupiony element
        Task.Delay(300).ContinueWith(_ =>
        {
            try
            {
                ReadFocusedElement();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"DialogMonitor: Błąd czytania focus: {ex.Message}");
            }
        });
    }

    private void AnnounceWindow(IntPtr hwnd, string title)
    {
        if (!string.IsNullOrEmpty(title))
        {
            _speechManager.Speak($"Okno: {title}", interrupt: true);
        }

        // Odczekaj chwilę i przeczytaj skupiony element
        Task.Delay(200).ContinueWith(_ =>
        {
            try
            {
                ReadFocusedElement();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"DialogMonitor: Błąd czytania focus: {ex.Message}");
            }
        });
    }

    private void ReadFocusedElement()
    {
        try
        {
            var focusedElement = AutomationElement.FocusedElement;
            if (focusedElement != null)
            {
                string description = UIAutomationHelper.GetElementDescription(focusedElement);
                if (!string.IsNullOrEmpty(description))
                {
                    _speechManager.Speak(description, interrupt: false);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"DialogMonitor: Błąd odczytu elementu: {ex.Message}");
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        if (_hookHandle != IntPtr.Zero)
        {
            UnhookWinEvent(_hookHandle);
            _hookHandle = IntPtr.Zero;
            Console.WriteLine("DialogMonitor: Hook usunięty");
        }

        _disposed = true;
    }
}
