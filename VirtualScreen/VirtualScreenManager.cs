using System.Runtime.InteropServices;
using System.Windows.Automation;
using ScreenReader.InputGestures;
using ScreenReader.Settings;
using ScreenReader.Speech;

namespace ScreenReader.VirtualScreen;

/// <summary>
/// Stan gestu - state machine zgodnie z Apple UIGestureRecognizer
/// Źródło: https://developer.apple.com/documentation/uikit/touches_presses_and_gestures/implementing_a_custom_gesture_recognizer/about_the_gesture_recognizer_state_machine
/// </summary>
internal enum GestureState
{
    /// <summary>Brak aktywnego gestu</summary>
    Idle,
    /// <summary>Gest możliwy - palec dotknął ekranu</summary>
    Possible,
    /// <summary>Gest rozpoczęty - potwierdzone minimum ruchu</summary>
    Began,
    /// <summary>Gest w trakcie - ciągły ruch</summary>
    Changed,
    /// <summary>Gest zakończony pomyślnie</summary>
    Ended,
    /// <summary>Gest anulowany (np. timeout)</summary>
    Cancelled,
    /// <summary>Gest nie rozpoznany</summary>
    Failed
}

/// <summary>
/// Zarządza wirtualnym ekranem dotykowym - eksploracja systemu przez touchpad/myszkę.
///
/// Gdy włączony:
/// - Myszka/touchpad działają jak symulowany ekran dotykowy
/// - Eksploracja działa globalnie (nie tylko w jednym oknie)
/// - Wolny ruch = eksploracja (element pod kursorem)
/// - Szybki ruch (swipe) lewo/prawo = poprzedni/następny element (jak NumPad 4/6)
/// - Szybki ruch (swipe) góra/dół = rodzic/dziecko (jak NumPad 8/2)
/// - Zatrzymanie ruchu = koniec gestu, wykrywa tap (odczytaj) i double-tap (aktywuj)
/// - Automatyczne przełączanie kontekstu na aktywne okno (cicho, bez dźwięków)
///
/// Logika ekranu dotykowego:
/// - Przerwa w ruchu > 150ms = palec oderwany, nowy gest
/// - Swipe = ruch >30px w <300ms
/// - Tap = ruch <20px w <300ms
/// - Przechwytuje wszystkie zdarzenia myszy/touchpada gdy włączony
/// </summary>
public class VirtualScreenManager : IDisposable
{
    private readonly SpeechManager _speechManager;
    private readonly SoundManager _soundManager;
    private readonly DialManager _dialManager;

    // Maksymalizowane okno dla Virtual Screen
    private IntPtr _maximizedWindow = IntPtr.Zero;
    private bool _wasMaximized = false;

    // Monitorowanie zmiany okna
    private IntPtr _windowEventHook = IntPtr.Zero;
    private WinEventDelegate? _windowEventDelegate;

    private bool _isEnabled;
    private bool _disposed;

    // UWAGA: ScreenReaderEngine.Instance?.CurrentElement został przeniesiony do ScreenReaderEngine.Instance.CurrentElement
    // dla synchronizacji z nawigacją obiektową (NumPad)

    // Stan gestu dotykowego - state machine zgodnie z Apple UIGestureRecognizer
    private GestureState _gestureState = GestureState.Idle;
    private int _gestureStartX;
    private int _gestureStartY;
    private int _gestureCurrentX;
    private int _gestureCurrentY;
    private int _gesturePreviousX;
    private int _gesturePreviousY;
    private DateTime _gestureStartTime;
    private DateTime _lastMoveTime;
    private bool _explorationStarted = false; // Czy eksploracja się rozpoczęła (po 500ms delay)

    // Velocity tracking dla swipe detection
    private float _velocityX;  // px/ms
    private float _velocityY;  // px/ms
    private const int VelocitySamples = 5;
    private readonly Queue<(DateTime time, int x, int y)> _velocityQueue = new();

    // Eksploracja przez dotyk
    private DateTime _lastExploreTime;

    // Tap detection
    private DateTime _lastTapTime;
    private int _lastTapX;
    private int _lastTapY;

    // Parametry gestów - zgodnie z best practices z badań
    // Źródła: TinyGesture, ZingTouch, @use-gesture, Android GestureDetector
    private const int TapThreshold = 3;         // 3px - standard z @use-gesture (tapsThreshold)
    private const int PressThreshold = 8;       // 8px - z TinyGesture (pressThreshold)
    private const float MinSwipeVelocity = 0.2f; // 0.2 px/ms - minimum dla swipe (zmniejszone dla łatwiejszych swipów)
    private const int MinSwipeDistance = 30;    // 30px - minimalna odległość dla swipe
    private const int SwipeMaxTimeMs = 300;     // 300ms - grace period dla gestu (zgodnie z badaniami)
    private const int GestureTimeoutMs = 150;   // 150ms - timeout = koniec gestu (palec oderwany) PO rozpoczęciu eksploracji
    private const int ExplorationStartDelayMs = 500; // 500ms - delay przed rozpoczęciem eksploracji (żeby swipey się dobrze wykrywały)
    private const int DoubleTapTimeMs = 500;    // 500ms - zgodnie z accessibility best practices (minimum)
    private const int ExploreDebounceMs = 50;   // 50ms - debounce dla eksploracji

    // Double-click (dla myszy) - również 500ms zgodnie z accessibility
    private DateTime _lastClickTime = DateTime.MinValue;
    private const int DoubleClickMs = 500;

    // Hook myszy
    private IntPtr _mouseHookId = IntPtr.Zero;
    private LowLevelMouseProc? _mouseProc;
    private short _lastWheelDelta;

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

    #region P/Invoke

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern bool ClipCursor(IntPtr lpRect);

    [DllImport("user32.dll")]
    private static extern bool ClipCursor(ref RECT lpRect);

    [DllImport("user32.dll")]
    private static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

    [DllImport("user32.dll")]
    private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

    [DllImport("user32.dll")]
    private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

    private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

    private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
    private const uint WINEVENT_OUTOFCONTEXT = 0x0000;

    private const int SW_MAXIMIZE = 3;
    private const int SW_RESTORE = 9;

    // SendInput structures
    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;
        public MOUSEINPUT mi;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    private const uint INPUT_MOUSE = 0;
    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;
    private const uint MOUSEEVENTF_ABSOLUTE = 0x8000;
    private const uint MOUSEEVENTF_MOVE = 0x0001;

    private const int WH_MOUSE_LL = 14;
    private const int WM_MOUSEMOVE = 0x0200;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_LBUTTONUP = 0x0202;
    private const int WM_RBUTTONDOWN = 0x0204;
    private const int WM_RBUTTONUP = 0x0205;
    private const int WM_MBUTTONDOWN = 0x0207;
    private const int WM_MBUTTONUP = 0x0208;
    private const int WM_MOUSEWHEEL = 0x020A;

    private const byte VK_APPS = 0x5D;
    private const byte VK_LWIN = 0x5B;
    private const byte VK_TAB = 0x09;
    private const byte VK_MENU = 0x12;      // Alt
    private const byte VK_SHIFT = 0x10;
    private const byte VK_PRIOR = 0x21;     // Page Up
    private const byte VK_NEXT = 0x22;      // Page Down
    private const byte VK_D = 0x44;
    private const uint KEYEVENTF_KEYUP = 0x0002;

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WINDOWPLACEMENT
    {
        public int length;
        public int flags;
        public int showCmd;
        public POINT ptMinPosition;
        public POINT ptMaxPosition;
        public RECT rcNormalPosition;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public uint mouseData;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    #endregion

    public bool IsEnabled => _isEnabled;
    public event Action<bool>? EnabledChanged;
    public event Action<AutomationElement?, float>? ElementChanged;

    public VirtualScreenManager(SpeechManager speechManager, SoundManager soundManager, DialManager dialManager)
    {
        _speechManager = speechManager;
        _soundManager = soundManager;
        _dialManager = dialManager;
    }

    /// <summary>
    /// Włącza eksplorację myszą - mysz działa jak palec na ekranie dotykowym.
    /// Działa globalnie na całym pulpicie i wszystkich oknach.
    /// Monitoruje zmianę okna dla automatycznej synchronizacji kontekstu.
    /// </summary>
    public void Enable()
    {
        if (_isEnabled) return;

        _isEnabled = true;

        // Ustaw początkowy element (bieżący fokus)
        ScreenReaderEngine.Instance?.CurrentElement = UIAutomationHelper.GetFocusedElement();

        // Zainstaluj hook myszy
        InstallMouseHook();

        // Zainstaluj hook zmiany okna (dla synchronizacji kontekstu, nie dla maksymalizacji)
        InstallWindowEventHook();

        _soundManager.PlayVirtualScreenOn();
        _speechManager.Speak("Wirtualny ekran włączony");

        EnabledChanged?.Invoke(true);
        Console.WriteLine("VirtualScreen: Włączono eksplorację myszą - działa globalnie na pulpicie");
    }

    /// <summary>
    /// Wyłącza eksplorację myszą - przywraca normalne działanie
    /// </summary>
    public void Disable()
    {
        if (!_isEnabled) return;

        _isEnabled = false;
        UninstallMouseHook();
        UninstallWindowEventHook();

        _soundManager.PlayVirtualScreenOff();
        _speechManager.Speak("Wirtualny ekran wyłączony");

        ScreenReaderEngine.Instance?.CurrentElement = null;

        EnabledChanged?.Invoke(false);
        Console.WriteLine("VirtualScreen: Wyłączono eksplorację myszą");
    }

    public void Toggle()
    {
        if (_isEnabled) Disable();
        else Enable();
    }

    /// <summary>
    /// Maksymalizuje aktywne okno i ogranicza kursor do jego granic.
    /// Zapamiętuje poprzedni stan okna aby móc go przywrócić.
    /// </summary>
    private void MaximizeAndClipToWindow()
    {
        try
        {
            _maximizedWindow = GetForegroundWindow();
            if (_maximizedWindow == IntPtr.Zero)
            {
                Console.WriteLine("VirtualScreen: Brak aktywnego okna");
                return;
            }

            // Sprawdź czy okno jest już zmaksymalizowane
            WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf(placement);
            if (GetWindowPlacement(_maximizedWindow, ref placement))
            {
                _wasMaximized = (placement.showCmd == SW_MAXIMIZE);

                if (!_wasMaximized)
                {
                    // Maksymalizuj okno
                    ShowWindow(_maximizedWindow, SW_MAXIMIZE);
                    Console.WriteLine("VirtualScreen: Zmaksymalizowano okno");
                    System.Threading.Thread.Sleep(200); // Poczekaj na animację
                }
                else
                {
                    Console.WriteLine("VirtualScreen: Okno już zmaksymalizowane");
                }
            }

            // Ogranicz kursor do granic okna
            if (GetWindowRect(_maximizedWindow, out RECT rect))
            {
                ClipCursor(ref rect);
                Console.WriteLine($"VirtualScreen: Ograniczono kursor do ({rect.Left},{rect.Top})-({rect.Right},{rect.Bottom})");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"VirtualScreen.MaximizeAndClip ERROR: {ex.Message}");
        }
    }

    /// <summary>
    /// Przywraca okno do poprzedniego stanu i usuwa ograniczenie kursora.
    /// </summary>
    private void RestoreWindowAndUnclip()
    {
        try
        {
            // Usuń ograniczenie kursora
            ClipCursor(IntPtr.Zero);
            Console.WriteLine("VirtualScreen: Usunięto ograniczenie kursora");

            // Przywróć okno jeśli było niezmaksymalizowane
            if (_maximizedWindow != IntPtr.Zero && !_wasMaximized)
            {
                ShowWindow(_maximizedWindow, SW_RESTORE);
                Console.WriteLine("VirtualScreen: Przywrócono poprzedni rozmiar okna");
            }

            _maximizedWindow = IntPtr.Zero;
            _wasMaximized = false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"VirtualScreen.RestoreAndUnclip ERROR: {ex.Message}");
        }
    }

    #region Window Event Hook

    /// <summary>
    /// Instaluje hook monitorujący zmianę aktywnego okna (EVENT_SYSTEM_FOREGROUND).
    /// </summary>
    private void InstallWindowEventHook()
    {
        _windowEventDelegate = WindowEventCallback;
        _windowEventHook = SetWinEventHook(
            EVENT_SYSTEM_FOREGROUND,
            EVENT_SYSTEM_FOREGROUND,
            IntPtr.Zero,
            _windowEventDelegate,
            0,
            0,
            WINEVENT_OUTOFCONTEXT
        );

        if (_windowEventHook != IntPtr.Zero)
        {
            Console.WriteLine("VirtualScreen: Zainstalowano hook monitorowania zmiany okna");
        }
        else
        {
            Console.WriteLine("VirtualScreen: Błąd instalacji hooka monitorowania okna");
        }
    }

    /// <summary>
    /// Odinstalowuje hook monitorowania zmiany okna.
    /// </summary>
    private void UninstallWindowEventHook()
    {
        if (_windowEventHook != IntPtr.Zero)
        {
            UnhookWinEvent(_windowEventHook);
            _windowEventHook = IntPtr.Zero;
            Console.WriteLine("VirtualScreen: Odinstalowano hook monitorowania okna");
        }
    }

    /// <summary>
    /// Callback wywoływany gdy zmienia się aktywne okno.
    /// Automatycznie synchronizuje kontekst Virtual Screen z nowym oknem.
    /// </summary>
    private void WindowEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (!_isEnabled || eventType != EVENT_SYSTEM_FOREGROUND)
            return;

        if (hwnd == IntPtr.Zero)
            return;

        try
        {
            Console.WriteLine($"VirtualScreen: Wykryto zmianę okna (hwnd={hwnd})");

            // Zaktualizuj bieżący element na fokus w nowym oknie (synchronizacja kontekstu)
            ScreenReaderEngine.Instance?.CurrentElement = UIAutomationHelper.GetFocusedElement();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"VirtualScreen.WindowEventCallback ERROR: {ex.Message}");
        }
    }

    #endregion

    #region Mouse Hook

    private void InstallMouseHook()
    {
        _mouseProc = MouseHookCallback;
        using var curProcess = System.Diagnostics.Process.GetCurrentProcess();
        using var curModule = curProcess.MainModule!;
        _mouseHookId = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, GetModuleHandle(curModule.ModuleName), 0);

        if (_mouseHookId == IntPtr.Zero)
        {
            Console.WriteLine("VirtualScreen: Błąd instalacji hooka");
        }
    }

    private void UninstallMouseHook()
    {
        if (_mouseHookId != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_mouseHookId);
            _mouseHookId = IntPtr.Zero;
        }
    }

    /// <summary>
    /// Hook myszy - PRZECHWYTUJE WSZYSTKO gdy włączony
    /// </summary>
    private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && _isEnabled)
        {
            var hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
            int x = hookStruct.pt.X;
            int y = hookStruct.pt.Y;
            int message = wParam.ToInt32();

            if (message == WM_MOUSEWHEEL)
            {
                _lastWheelDelta = (short)((hookStruct.mouseData >> 16) & 0xFFFF);
            }

            // Przetwórz i ZABLOKUJ wszystkie zdarzenia myszy
            ProcessMouseEvent(message, x, y);
            return (IntPtr)1; // Blokuj przekazanie do systemu
        }

        return CallNextHookEx(_mouseHookId, nCode, wParam, lParam);
    }

    #endregion

    #region Event Processing - Symulowany ekran dotykowy (myszka/touchpad)

    private void ProcessMouseEvent(int message, int x, int y)
    {
        switch (message)
        {
            case WM_MOUSEMOVE:
                // Ruch myszki/touchpada = symulowany ekran dotykowy
                // Wolny ruch = eksploracja, szybki ruch = swipe
                HandleTouchpadMove(x, y);
                break;

            case WM_LBUTTONDOWN:
                // Klik lewym = natychmiastowy tap (dla myszy)
                HandleMouseClick(x, y, isDouble: false);
                break;

            case WM_RBUTTONDOWN:
                // Prawy przycisk = menu kontekstowe
                ExecuteContextMenu();
                break;

            case WM_MBUTTONDOWN:
                // Środkowy przycisk = menu Start
                ExecuteStartMenu();
                break;

            case WM_MOUSEWHEEL:
                // Scroll = zmiana kategorii pokrętła
                if (_lastWheelDelta > 0)
                    NavigateDialPrevious();
                else
                    NavigateDialNext();
                break;
        }
    }

    /// <summary>
    /// Główna obsługa ruchu - symuluje ekran dotykowy dla myszki/touchpada.
    /// Implementuje state machine zgodnie z Apple UIGestureRecognizer.
    ///
    /// Myszka/TouchPad działają jak ekran dotykowy:
    /// - Eksploracja rozpoczyna się DOPIERO po 500ms przytrzymania palca (żeby swipey się dobrze wykrywały)
    /// - Szybki ruch (swipe >30px + velocity >0.2 px/ms) wykrywany w ciągu 300ms = nawigacja po elementach
    /// - Po rozpoczęciu eksploracji, kontynuuje się podczas ruchu
    /// - Przerwa w ruchu >150ms PO rozpoczęciu eksploracji = koniec gestu
    /// - Małe przesunięcie (<3px) + krótki czas = tap/double-tap
    ///
    /// Źródła:
    /// - TinyGesture: velocity threshold 10 px/ms (dostosowane do 0.2 dla touchpada)
    /// - @use-gesture: tap threshold 3px
    /// - Accessibility research: 500ms double-tap, 300ms grace period
    /// </summary>
    private void HandleTouchpadMove(int x, int y)
    {
        var now = DateTime.Now;
        var timeSinceLastMove = (now - _lastMoveTime).TotalMilliseconds;

        // State machine: sprawdź timeout (palec oderwany) - tylko jeśli eksploracja się rozpoczęła
        if (_explorationStarted && timeSinceLastMove > GestureTimeoutMs)
        {
            // Timeout = koniec gestu (tylko po rozpoczęciu eksploracji)
            if (_gestureState != GestureState.Idle)
            {
                EndGesture(_gestureCurrentX, _gestureCurrentY);
            }

            // Rozpocznij nowy gest
            BeginGesture(x, y);
        }
        else if (!_explorationStarted && timeSinceLastMove > SwipeMaxTimeMs)
        {
            // Jeśli brak ruchu przez 300ms PRZED rozpoczęciem eksploracji, zakończ gest
            // (swipe nie został wykryty, a palec jest nieruchomy)
            if (_gestureState != GestureState.Idle)
            {
                EndGesture(_gestureCurrentX, _gestureCurrentY);
            }
            BeginGesture(x, y);
        }

        // Update pozycji
        _gesturePreviousX = _gestureCurrentX;
        _gesturePreviousY = _gestureCurrentY;
        _gestureCurrentX = x;
        _gestureCurrentY = y;
        _lastMoveTime = now;

        // Update velocity tracking
        UpdateVelocity(now, x, y);

        // State machine transitions
        switch (_gestureState)
        {
            case GestureState.Idle:
                // Nie powinno się zdarzyć - BeginGesture już ustawił Possible
                BeginGesture(x, y);
                break;

            case GestureState.Possible:
                // Sprawdź czy ruch przekroczył próg PressThreshold (potwierdzenie gestu)
                int dx = x - _gestureStartX;
                int dy = y - _gestureStartY;
                int totalMovement = Math.Abs(dx) + Math.Abs(dy);

                if (totalMovement > PressThreshold)
                {
                    // Ruch potwierdzony - przejdź do Began
                    _gestureState = GestureState.Began;
                    GestureBegan();
                }
                else
                {
                    // Mały ruch - sprawdź czy minęło 500ms aby rozpocząć eksplorację
                    var gestureElapsed = (now - _gestureStartTime).TotalMilliseconds;
                    if (gestureElapsed >= ExplorationStartDelayMs && !_explorationStarted)
                    {
                        // Rozpocznij eksplorację po 500ms przytrzymania
                        _explorationStarted = true;
                        Console.WriteLine("Exploration: Started (po 500ms przytrzymania)");
                    }

                    // Wykonuj eksplorację tylko jeśli się rozpoczęła
                    if (_explorationStarted)
                    {
                        DoExploration(x, y);
                    }
                }
                break;

            case GestureState.Began:
            case GestureState.Changed:
                // Gest w trakcie
                _gestureState = GestureState.Changed;
                GestureChanged(x, y);
                break;

            case GestureState.Ended:
            case GestureState.Cancelled:
            case GestureState.Failed:
                // Gest zakończony - nie powinno się zdarzyć
                BeginGesture(x, y);
                break;
        }
    }

    /// <summary>
    /// Update velocity tracking z kolejki ostatnich pozycji
    /// Używa sliding window dla obliczenia średniej prędkości
    /// </summary>
    private void UpdateVelocity(DateTime now, int x, int y)
    {
        // Dodaj do kolejki
        _velocityQueue.Enqueue((now, x, y));

        // Ogranicz kolejkę do VelocitySamples
        while (_velocityQueue.Count > VelocitySamples)
            _velocityQueue.Dequeue();

        // Oblicz velocity na podstawie pierwszego i ostatniego sampla
        if (_velocityQueue.Count >= 2)
        {
            var first = _velocityQueue.First();
            var last = _velocityQueue.Last();

            var deltaTime = (last.time - first.time).TotalMilliseconds;
            if (deltaTime > 0)
            {
                _velocityX = (last.x - first.x) / (float)deltaTime;  // px/ms
                _velocityY = (last.y - first.y) / (float)deltaTime;  // px/ms
            }
        }
    }

    /// <summary>
    /// Gest rozpoczęty - przeszedł próg PressThreshold
    /// </summary>
    private void GestureBegan()
    {
        Console.WriteLine("Gesture: Began");
        // Można dodać dźwięk lub feedback
    }

    /// <summary>
    /// Gest w trakcie - ciągły ruch
    /// Sprawdza czy to swipe lub eksploracja
    /// </summary>
    private void GestureChanged(int x, int y)
    {
        var now = DateTime.Now;
        var gestureElapsed = (now - _gestureStartTime).TotalMilliseconds;

        // Sprawdź czy to swipe (velocity + distance)
        int dx = x - _gestureStartX;
        int dy = y - _gestureStartY;
        int totalDistance = Math.Abs(dx) + Math.Abs(dy);

        // Oblicz całkowitą velocity
        float totalVelocity = MathF.Sqrt(_velocityX * _velocityX + _velocityY * _velocityY);

        // Swipe detection: distance > threshold AND velocity > threshold (w ciągu 300ms)
        if (gestureElapsed < SwipeMaxTimeMs &&
            totalDistance > MinSwipeDistance &&
            totalVelocity > MinSwipeVelocity)
        {
            // To jest swipe!
            ExecuteSwipe(dx, dy, Math.Abs(dx), Math.Abs(dy));

            // Zakończ gest po wykonaniu swipe
            _gestureState = GestureState.Ended;
            return;
        }

        // Sprawdź czy rozpocząć eksplorację (po 500ms wolnego ruchu)
        if (!_explorationStarted && gestureElapsed >= ExplorationStartDelayMs)
        {
            _explorationStarted = true;
            Console.WriteLine("Exploration: Started during slow movement (po 500ms)");
        }

        // Kontynuuj eksplorację jeśli się rozpoczęła
        if (_explorationStarted)
        {
            DoExploration(x, y);
        }
    }

    /// <summary>
    /// Rozpoczyna nowy gest - state machine: Idle -> Possible
    /// </summary>
    private void BeginGesture(int x, int y)
    {
        _gestureState = GestureState.Possible;
        _gestureStartX = x;
        _gestureStartY = y;
        _gestureCurrentX = x;
        _gestureCurrentY = y;
        _gesturePreviousX = x;
        _gesturePreviousY = y;
        _gestureStartTime = DateTime.Now;
        _lastMoveTime = DateTime.Now;
        _explorationStarted = false; // Reset flagi eksploracji

        // Wyczyść velocity tracking
        _velocityQueue.Clear();
        _velocityX = 0;
        _velocityY = 0;

        Console.WriteLine($"Gesture: Began at ({x},{y})");
    }

    /// <summary>
    /// Kończy gest - state machine: Any -> Ended/Cancelled
    /// Sprawdza czy to był tap
    /// </summary>
    private void EndGesture(int x, int y)
    {
        if (_gestureState == GestureState.Idle)
            return;

        var now = DateTime.Now;
        var gestureElapsed = (now - _gestureStartTime).TotalMilliseconds;

        // Sprawdź czy to tap (mały ruch, krótki czas)
        int dx = x - _gestureStartX;
        int dy = y - _gestureStartY;
        int totalMovement = Math.Abs(dx) + Math.Abs(dy);

        // Tap detection: mały ruch (<TapThreshold) + krótki czas (<SwipeMaxTimeMs)
        if (totalMovement < TapThreshold && gestureElapsed < SwipeMaxTimeMs)
        {
            DetectTap(x, y);
            _gestureState = GestureState.Ended;
        }
        else if (_gestureState == GestureState.Began || _gestureState == GestureState.Changed)
        {
            // Gest zakończony normalnie (nie był tap ani swipe)
            _gestureState = GestureState.Ended;
        }
        else
        {
            // Gest nie rozpoznany lub anulowany
            _gestureState = GestureState.Cancelled;
        }

        Console.WriteLine($"Gesture: Ended (state={_gestureState}, movement={totalMovement}px, time={gestureElapsed:F0}ms)");

        // Reset do Idle
        _gestureState = GestureState.Idle;
    }

    /// <summary>
    /// Wykrywa single/double tap
    /// Zgodnie z accessibility best practices: 500ms na double-tap
    /// Źródło: TapNav research (arxiv.org/html/2510.14267)
    /// </summary>
    private void DetectTap(int x, int y)
    {
        var now = DateTime.Now;
        var timeSinceLastTap = (now - _lastTapTime).TotalMilliseconds;

        // Sprawdź double-tap (dwa tapy blisko siebie w czasie <500ms)
        if (timeSinceLastTap < DoubleTapTimeMs &&
            Math.Abs(x - _lastTapX) < TapThreshold * 2 &&
            Math.Abs(y - _lastTapY) < TapThreshold * 2)
        {
            // Double-tap = aktywuj element
            ActivateElement();
            Console.WriteLine($"Gesture: Double-tap at ({x},{y})");
            return;
        }

        // Single tap
        _lastTapTime = now;
        _lastTapX = x;
        _lastTapY = y;

        // Odczytaj element (single tap)
        ReadCurrentElement();
        Console.WriteLine($"Gesture: Single tap at ({x},{y})");
    }

    /// <summary>
    /// Obsługa kliknięcia myszą (dla użytkowników myszy)
    /// </summary>
    private void HandleMouseClick(int x, int y, bool isDouble)
    {
        var now = DateTime.Now;

        if ((now - _lastClickTime).TotalMilliseconds < DoubleClickMs)
        {
            // Double-click
            ActivateElement();
            _lastClickTime = DateTime.MinValue;
        }
        else
        {
            // Single click
            _lastClickTime = now;
            ReadCurrentElement();
        }
    }

    /// <summary>
    /// Wykonuje swipe - działa jak nawigacja obiektowa NumPad
    /// Synchronizuje się z nawigacją NumPad przez współdzielony ScreenReaderEngine.Instance?.CurrentElement
    ///
    /// Kierunki:
    /// - Poziomo: lewo/prawo = poprzedni/następny sibling (NumPad 4/6)
    /// - Pionowo: góra/dół = parent/child (NumPad 8/2)
    ///
    /// Velocity detection zgodnie z ZingTouch (escape velocity)
    /// </summary>
    private void ExecuteSwipe(int dx, int dy, int absDx, int absDy)
    {
        // Określ kierunek na podstawie większej składowej
        if (absDx > absDy)
        {
            // Swipe poziomy
            if (dx > 0)
            {
                NavigateNextSibling();  // Swipe prawo = następny sibling (NumPad 6)
                Console.WriteLine($"Swipe: prawo (dx={dx}, velocity={_velocityX:F2} px/ms)");
            }
            else
            {
                NavigatePreviousSibling();  // Swipe lewo = poprzedni sibling (NumPad 4)
                Console.WriteLine($"Swipe: lewo (dx={dx}, velocity={_velocityX:F2} px/ms)");
            }
        }
        else
        {
            // Swipe pionowy
            if (dy > 0)
            {
                // Swipe dół = pierwsze dziecko (NumPad 2)
                NavigateFirstChild();
                Console.WriteLine($"Swipe: dół (dy={dy}, velocity={_velocityY:F2} px/ms)");
            }
            else
            {
                // Swipe góra = rodzic (NumPad 8)
                NavigateParent();
                Console.WriteLine($"Swipe: góra (dy={dy}, velocity={_velocityY:F2} px/ms)");
            }
        }
    }

    /// <summary>
    /// Eksploracja - znajdź element BEZPOŚREDNIO pod pozycją ekranową.
    /// Używa AutomationElement.FromPoint() - działa globalnie (nie tylko w jednym oknie).
    /// Wywoływane przez HandleTouchpadMove (symulowany ekran dotykowy) i OnMultiTouchExplore (prawdziwy touchpad).
    /// </summary>
    private void DoExploration(int x, int y)
    {
        var now = DateTime.Now;
        if ((now - _lastExploreTime).TotalMilliseconds < ExploreDebounceMs)
            return;

        _lastExploreTime = now;

        Console.WriteLine($"DoExploration: x={x}, y={y}");

        try
        {
            // Użyj zaawansowanego detektora - rekursywne skanowanie dla małych elementów
            var element = UIAutomation.ElementDetector.FindElementAtPoint(x, y, searchRadius: 10);

            if (element == null)
            {
                Console.WriteLine("DoExploration: element is NULL");
                return;
            }

            // Sprawdź czy to ten sam element
            if (ScreenReaderEngine.Instance?.CurrentElement != null)
            {
                try
                {
                    if (Automation.Compare(element, ScreenReaderEngine.Instance.CurrentElement))
                    {
                        Console.WriteLine("DoExploration: Ten sam element - pomijam");
                        return; // Ten sam element - nie ogłaszaj ponownie
                    }
                }
                catch { }
            }

            // Synchronizuj z ScreenReaderEngine
            if (ScreenReaderEngine.Instance != null)
            {
                ScreenReaderEngine.Instance.CurrentElement = element;
            }

            // Oblicz pozycję przestrzenną dla 3D audio
            CalculateAndStoreSpatialPosition(x, y);

            // Pobierz obliczoną pozycję
            var (azimuth, elevation) = ScreenReaderEngine.Instance?.CurrentElementPosition ?? (0f, 0f);

            Console.WriteLine($"DoExploration: Znaleziono element - ogłaszam z 3D audio (azimuth={azimuth * 180 / MathF.PI:F0}°, elevation={elevation * 180 / MathF.PI:F0}°)");

            // Odtwórz dźwięk i mowę z pozycjonowaniem 3D
            PlayNavigationSound(element, azimuth, elevation);
            AnnounceElement(element, azimuth, elevation);

            ElementChanged?.Invoke(element, 0.5f);  // Bez listy nie ma pozycji procentowej
        }
        catch (Exception ex)
        {
            Console.WriteLine($"DoExploration ERROR: {ex.Message}");
        }
    }

    /// <summary>
    /// Menu kontekstowe (2 palce tap lub prawy przycisk)
    /// </summary>
    private void ExecuteContextMenu()
    {
        _soundManager.PlaySystemItem();
        keybd_event(VK_APPS, 0, 0, UIntPtr.Zero);
        keybd_event(VK_APPS, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        Console.WriteLine("Gest: Menu kontekstowe");
    }

    /// <summary>
    /// Menu Start (3 palce tap lub środkowy przycisk)
    /// </summary>
    private void ExecuteStartMenu()
    {
        _soundManager.PlaySystemItem();
        keybd_event(VK_LWIN, 0, 0, UIntPtr.Zero);
        keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        Console.WriteLine("Gest: Menu Start");
    }

    #endregion

    #region Navigation (jak NumPad - nawigacja obiektowa)

    /// <summary>
    /// Upewnia się że nawigacja działa w kontekście aktywnego okna.
    /// Jeśli użytkownik przełączył okno (np. Alt+Tab), aktualizuje ScreenReaderEngine.Instance?.CurrentElement.
    /// </summary>
    private void EnsureCurrentWindowContext()
    {
        try
        {
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero)
                return;

            var activeWindow = AutomationElement.FromHandle(foregroundWindow);
            if (activeWindow == null)
                return;

            // Sprawdź czy ScreenReaderEngine.Instance?.CurrentElement jest w aktywnym oknie
            if (ScreenReaderEngine.Instance?.CurrentElement != null)
            {
                try
                {
                    // Sprawdź czy element nadal istnieje i jest w aktywnym oknie
                    var elementWindow = GetElementWindow(ScreenReaderEngine.Instance?.CurrentElement);
                    if (elementWindow != null && Automation.Compare(elementWindow, activeWindow))
                    {
                        // Ten sam kontekst - OK
                        return;
                    }
                }
                catch
                {
                    // Element może nie istnieć - przełącz na nowe okno
                }
            }

            // Przełącz na aktywne okno (cicho - nie ogłaszaj, żeby nie przeszkadzać w nawigacji)
            Console.WriteLine("VirtualScreen: Przełączam kontekst na aktywne okno");

            // Spróbuj uzyskać element z fokusem w nowym oknie
            var focusedElement = UIAutomationHelper.GetFocusedElement();
            if (focusedElement != null)
            {
                ScreenReaderEngine.Instance?.CurrentElement = focusedElement;
            }
            else
            {
                // Fallback - użyj samego okna
                ScreenReaderEngine.Instance?.CurrentElement = activeWindow;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"VirtualScreen.EnsureCurrentWindowContext ERROR: {ex.Message}");
        }
    }

    /// <summary>
    /// Pobiera okno nadrzędne dla danego elementu
    /// </summary>
    private AutomationElement? GetElementWindow(AutomationElement element)
    {
        try
        {
            var walker = TreeWalker.ControlViewWalker;
            var current = element;

            while (current != null && !Automation.Compare(current, AutomationElement.RootElement))
            {
                if (current.Current.ControlType == ControlType.Window)
                    return current;

                current = walker.GetParent(current);
            }
        }
        catch { }

        return null;
    }

    /// <summary>
    /// Następny sibling (NumPad 6 / Swipe prawo)
    /// </summary>
    private void NavigateNextSibling()
    {
        EnsureCurrentWindowContext();

        if (ScreenReaderEngine.Instance?.CurrentElement == null)
        {
            _soundManager.PlayEdge();
            return;
        }

        try
        {
            var walker = TreeWalker.ControlViewWalker;
            var next = walker.GetNextSibling(ScreenReaderEngine.Instance?.CurrentElement);

            if (next != null)
            {
                ScreenReaderEngine.Instance?.CurrentElement = next;
                _soundManager.PlayCursor();
                AnnounceElement(next);
            }
            else
            {
                _soundManager.PlayEdge();
            }
        }
        catch
        {
            _soundManager.PlayEdge();
        }
    }

    /// <summary>
    /// Poprzedni sibling (NumPad 4 / Swipe lewo)
    /// </summary>
    private void NavigatePreviousSibling()
    {
        EnsureCurrentWindowContext();

        if (ScreenReaderEngine.Instance?.CurrentElement == null)
        {
            _soundManager.PlayEdge();
            return;
        }

        try
        {
            var walker = TreeWalker.ControlViewWalker;
            var prev = walker.GetPreviousSibling(ScreenReaderEngine.Instance?.CurrentElement);

            if (prev != null)
            {
                ScreenReaderEngine.Instance?.CurrentElement = prev;
                _soundManager.PlayCursor();
                AnnounceElement(prev);
            }
            else
            {
                _soundManager.PlayEdge();
            }
        }
        catch
        {
            _soundManager.PlayEdge();
        }
    }

    /// <summary>
    /// Rodzic (NumPad 8 / Swipe góra)
    /// </summary>
    private void NavigateParent()
    {
        EnsureCurrentWindowContext();

        if (ScreenReaderEngine.Instance?.CurrentElement == null)
        {
            _soundManager.PlayEdge();
            return;
        }

        try
        {
            var walker = TreeWalker.ControlViewWalker;
            var parent = walker.GetParent(ScreenReaderEngine.Instance?.CurrentElement);

            if (parent != null && !Automation.Compare(parent, AutomationElement.RootElement))
            {
                ScreenReaderEngine.Instance?.CurrentElement = parent;
                _soundManager.PlayZoomOut();
                AnnounceElement(parent);
            }
            else
            {
                _soundManager.PlayEdge();
            }
        }
        catch
        {
            _soundManager.PlayEdge();
        }
    }

    /// <summary>
    /// Pierwsze dziecko (NumPad 2 / Swipe dół)
    /// </summary>
    private void NavigateFirstChild()
    {
        EnsureCurrentWindowContext();

        if (ScreenReaderEngine.Instance?.CurrentElement == null)
        {
            _soundManager.PlayEdge();
            return;
        }

        try
        {
            var walker = TreeWalker.ControlViewWalker;
            var child = walker.GetFirstChild(ScreenReaderEngine.Instance?.CurrentElement);

            if (child != null)
            {
                ScreenReaderEngine.Instance?.CurrentElement = child;
                _soundManager.PlayZoomIn();
                AnnounceElement(child);
            }
            else
            {
                _soundManager.PlayEdge();
            }
        }
        catch
        {
            _soundManager.PlayEdge();
        }
    }

    /// <summary>
    /// Odczytaj bieżący element (NumPad 5 / Single tap)
    /// </summary>
    private void ReadCurrentElement()
    {
        if (ScreenReaderEngine.Instance?.CurrentElement == null)
        {
            _speechManager.Speak("Brak elementu");
            return;
        }

        PlayNavigationSound(ScreenReaderEngine.Instance?.CurrentElement);
        AnnounceElement(ScreenReaderEngine.Instance?.CurrentElement);
    }

    /// <summary>
    /// Aktywuj element (NumPad Enter / Double-tap / Double-click)
    /// ZSYNCHRONIZOWANE z nawigacją obiektową - używa ScreenReaderEngine.OnClickAction()
    /// </summary>
    private void ActivateElement()
    {
        if (ScreenReaderEngine.Instance?.CurrentElement == null)
        {
            _speechManager.Speak("Brak elementu");
            return;
        }

        // Dźwięk doubletab.ogg (jak double-click na touchpadzie)
        _soundManager.PlayDoubleClick();

        // Użyj unified activation logic z ScreenReaderEngine
        // To zapewnia identyczne zachowanie jak NumPad /
        ScreenReaderEngine.Instance.OnClickAction();
    }

    /// <summary>
    /// Oblicza pozycję przestrzenną elementu (azymut/elewacja) z współrzędnych ekranu
    /// Używane do 3D audio podczas eksploracji myszą
    /// </summary>
    private void CalculateAndStoreSpatialPosition(int x, int y)
    {
        // Pobierz wymiary ekranu
        int screenWidth = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;
        int screenHeight = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height;

        // Normalizuj do 0.0-1.0
        float panX = x / (float)screenWidth;
        float panY = y / (float)screenHeight;

        // Mapuj na azymut: -90° (lewo) do +90° (prawo)
        // -π/2 radians do +π/2 radians
        float azimuth = (panX - 0.5f) * MathF.PI;

        // Mapuj na elewację: +45° (góra) do -45° (dół)
        // +π/4 radians do -π/4 radians
        float elevation = (0.5f - panY) * MathF.PI / 2;

        // Zapisz w ScreenReaderEngine dla późniejszego użycia
        if (ScreenReaderEngine.Instance != null)
        {
            ScreenReaderEngine.Instance.CurrentElementPosition = (azimuth, elevation);
        }
    }

    #endregion

    #region Helpers

    /// <summary>
    /// Ogłasza element za pomocą syntezatora mowy z opcjonalnym pozycjonowaniem 3D.
    /// TYLKO podczas eksploracji myszą przekazuj azimuth/elevation dla 3D audio.
    /// </summary>
    /// <param name="element">Element do ogłoszenia</param>
    /// <param name="azimuth">Kąt azymutalny (opcjonalny, null = normalne audio)</param>
    /// <param name="elevation">Kąt elewacji (opcjonalny, null = normalne audio)</param>
    private void AnnounceElement(AutomationElement element, float? azimuth = null, float? elevation = null)
    {
        try
        {
            var name = element.Current.Name ?? "";
            var controlType = element.Current.ControlType;
            var typeName = UIAutomationHelper.GetPolishControlType(controlType);

            string text = !string.IsNullOrWhiteSpace(name)
                ? $"{name}, {typeName}"
                : typeName;

            Console.WriteLine($"VirtualScreen.AnnounceElement: '{text}' (3D: {azimuth.HasValue})");

            // Jeśli azimuth/elevation podane - użyj 3D audio (tylko podczas eksploracji myszą)
            // W przeciwnym razie - normalne audio (swipey, NumPad)
            _speechManager.Speak(text, true, azimuth, elevation);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"VirtualScreen.AnnounceElement ERROR: {ex.Message}");
        }
    }


    private void PlayNavigationSound(AutomationElement element, float azimuth = 0f, float elevation = 0f)
    {
        try
        {
            var controlType = element.Current.ControlType;

            if (controlType == ControlType.ListItem)
            {
                // Dla elementów listy użyj pozycji względnej (jeśli dostępna)
                float pos = GetListItemPosition(element);
                _soundManager.PlayListItem(pos, azimuth, elevation);
            }
            else
            {
                // Dla pozostałych używamy cursor.ogg
                _soundManager.PlayCursor(azimuth, elevation);
            }
        }
        catch
        {
            _soundManager.PlayCursor(azimuth, elevation);
        }
    }

    /// <summary>
    /// Pobiera pozycję elementu listy (0.0-1.0) używając SelectionItemPattern
    /// </summary>
    private float GetListItemPosition(AutomationElement element)
    {
        try
        {
            // Spróbuj pobrać pozycję z SelectionItemPattern
            if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var pattern))
            {
                var container = ((SelectionItemPattern)pattern).Current.SelectionContainer;
                if (container != null)
                {
                    var walker = TreeWalker.ControlViewWalker;
                    var firstChild = walker.GetFirstChild(container);
                    int index = 0;
                    int total = 0;

                    var current = firstChild;
                    while (current != null)
                    {
                        if (current.Current.ControlType == ControlType.ListItem)
                        {
                            total++;
                            if (Automation.Compare(current, element))
                                index = total - 1;
                        }
                        current = walker.GetNextSibling(current);
                    }

                    if (total > 0)
                        return (float)index / Math.Max(1, total - 1);
                }
            }
        }
        catch { }

        return 0.5f; // Środek jeśli nie można określić pozycji
    }

    private void NavigateDialPrevious()
    {
        _soundManager.PlayDialItem();
        string? msg = _dialManager.ExecuteItemChange(false, _speechManager);
        if (!string.IsNullOrEmpty(msg))
            _speechManager.Speak(msg);
    }

    private void NavigateDialNext()
    {
        _soundManager.PlayDialItem();
        string? msg = _dialManager.ExecuteItemChange(true, _speechManager);
        if (!string.IsNullOrEmpty(msg))
            _speechManager.Speak(msg);
    }

    #endregion

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_isEnabled) Disable();
        UninstallMouseHook();
    }
}
