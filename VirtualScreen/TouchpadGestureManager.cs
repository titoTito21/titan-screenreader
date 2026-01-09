using System.Runtime.InteropServices;

namespace ScreenReader.VirtualScreen;

/// <summary>
/// Typy gestów touchpada zgodne z koncepcją wirtualnego ekranu.
/// </summary>
public enum TouchpadGesture
{
    None,

    // Gesty 1 palcem
    Explore,                    // Przesuwanie = eksploracja z 3D audio
    SwipeLeft,                  // Poprzedni element
    SwipeRight,                 // Następny element
    SwipeUp,                    // Poprzedni element pokrętła (dial)
    SwipeDown,                  // Następny element pokrętła (dial)
    SingleTap,                  // Odczytaj element
    DoubleTap,                  // Aktywuj element

    // Gesty 2 palcami
    TwoFingerTap,               // Menu kontekstowe
    TwoFingerDoubleTap,         // Pulpit
    TwoFingerSwipeLeft,         // Alt+Shift+Tab
    TwoFingerSwipeRight,        // Alt+Tab
    TwoFingerSwipeUp,           // Poprzednia kategoria pokrętła
    TwoFingerSwipeDown,         // Następna kategoria pokrętła

    // Gesty 3 palcami
    ThreeFingerTap,             // Menu Start
    ThreeFingerSwipeLeft,       // Page Up
    ThreeFingerSwipeRight,      // Page Down
}

/// <summary>
/// Zarządza gestami wielopalcowymi na Precision Touchpad.
/// Używa Raw Input API do wykrywania liczby palców i pozycji.
/// </summary>
public class TouchpadGestureManager : IDisposable
{
    private bool _disposed;
    private IntPtr _hwnd;

    // Stan dotyku
    private int _fingerCount;
    private readonly TouchPoint[] _fingers = new TouchPoint[5];
    private TouchPoint _gestureStart;
    private DateTime _touchStartTime;
    private DateTime _lastTapTime;
    private bool _gestureExecuted;

    // Parametry gestów
    private const int SwipeThreshold = 22;
    private const int TapThreshold = 12;
    private const int DoubleTapTimeMs = 350;
    private const int SwipeMaxTimeMs = 220;

    public event Action<TouchpadGesture, int, int>? GestureDetected;
    public event Action<int, int>? ExploreMove;

    #region P/Invoke dla Raw Input

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterRawInputDevices(RAWINPUTDEVICE[] pRawInputDevices, uint uiNumDevices, uint cbSize);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetRawInputData(IntPtr hRawInput, uint uiCommand, IntPtr pData, ref uint pcbSize, uint cbSizeHeader);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetRawInputDeviceList(IntPtr pRawInputDeviceList, ref uint puiNumDevices, uint cbSize);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetRawInputDeviceInfo(IntPtr hDevice, uint uiCommand, IntPtr pData, ref uint pcbSize);

    private const int WM_INPUT = 0x00FF;
    private const int RID_INPUT = 0x10000003;
    private const int RIM_TYPEHID = 2;
    private const int RIDEV_INPUTSINK = 0x00000100;

    // HID Usage Page i Usage dla Touchpad
    private const ushort HID_USAGE_PAGE_DIGITIZER = 0x0D;
    private const ushort HID_USAGE_DIGITIZER_TOUCH_PAD = 0x05;

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWINPUTDEVICE
    {
        public ushort usUsagePage;
        public ushort usUsage;
        public uint dwFlags;
        public IntPtr hwndTarget;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWINPUTHEADER
    {
        public uint dwType;
        public uint dwSize;
        public IntPtr hDevice;
        public IntPtr wParam;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWHID
    {
        public uint dwSizeHid;
        public uint dwCount;
        // bRawData follows
    }

    #endregion

    private struct TouchPoint
    {
        public int X;
        public int Y;
        public bool IsDown;
        public int Id;
    }

    public TouchpadGestureManager()
    {
        for (int i = 0; i < _fingers.Length; i++)
        {
            _fingers[i] = new TouchPoint();
        }
    }

    /// <summary>
    /// Rejestruje urządzenie touchpad dla Raw Input.
    /// </summary>
    public bool RegisterForRawInput(IntPtr hwnd)
    {
        _hwnd = hwnd;

        var devices = new RAWINPUTDEVICE[]
        {
            new RAWINPUTDEVICE
            {
                usUsagePage = HID_USAGE_PAGE_DIGITIZER,
                usUsage = HID_USAGE_DIGITIZER_TOUCH_PAD,
                dwFlags = RIDEV_INPUTSINK,
                hwndTarget = hwnd
            }
        };

        bool result = RegisterRawInputDevices(devices, 1, (uint)Marshal.SizeOf<RAWINPUTDEVICE>());

        if (result)
        {
            Console.WriteLine("TouchpadGesture: Zarejestrowano Raw Input dla touchpada");
        }
        else
        {
            int error = Marshal.GetLastWin32Error();
            Console.WriteLine($"TouchpadGesture: Błąd rejestracji Raw Input: {error}");
        }

        return result;
    }

    /// <summary>
    /// Przetwarza wiadomość WM_INPUT.
    /// </summary>
    public void ProcessRawInput(IntPtr lParam)
    {
        uint size = 0;
        GetRawInputData(lParam, RID_INPUT, IntPtr.Zero, ref size, (uint)Marshal.SizeOf<RAWINPUTHEADER>());

        if (size == 0)
            return;

        IntPtr buffer = Marshal.AllocHGlobal((int)size);
        try
        {
            if (GetRawInputData(lParam, RID_INPUT, buffer, ref size, (uint)Marshal.SizeOf<RAWINPUTHEADER>()) == size)
            {
                var header = Marshal.PtrToStructure<RAWINPUTHEADER>(buffer);

                if (header.dwType == RIM_TYPEHID)
                {
                    // Przetwórz dane HID touchpada
                    ProcessHidData(buffer, header.dwSize);
                }
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    /// <summary>
    /// Przetwarza dane HID z touchpada.
    /// Format zależy od konkretnego urządzenia - tu uproszczona wersja.
    /// </summary>
    private void ProcessHidData(IntPtr buffer, uint size)
    {
        // Dane HID są po RAWINPUTHEADER + RAWHID
        int headerSize = Marshal.SizeOf<RAWINPUTHEADER>();
        int hidHeaderSize = Marshal.SizeOf<RAWHID>();

        if (size < headerSize + hidHeaderSize)
            return;

        var hidHeader = Marshal.PtrToStructure<RAWHID>(buffer + headerSize);

        // Dane surowe są po nagłówku RAWHID
        IntPtr rawData = buffer + headerSize + hidHeaderSize;
        int dataSize = (int)(hidHeader.dwSizeHid * hidHeader.dwCount);

        // Tu trzeba parsować dane HID zgodnie z HID Report Descriptor
        // To jest skomplikowane i zależy od urządzenia
        // Na razie logujemy
        Console.WriteLine($"TouchpadGesture: HID data size={dataSize}, dwSizeHid={hidHeader.dwSizeHid}, dwCount={hidHeader.dwCount}");
    }

    /// <summary>
    /// Alternatywna metoda - używa WM_POINTER API (Windows 8+).
    /// Prostsze niż Raw Input HID.
    /// </summary>
    public void ProcessPointerInput(int pointerId, int x, int y, bool isDown, int fingerIndex)
    {
        if (fingerIndex < 0 || fingerIndex >= _fingers.Length)
            return;

        var wasDown = _fingers[fingerIndex].IsDown;

        _fingers[fingerIndex].X = x;
        _fingers[fingerIndex].Y = y;
        _fingers[fingerIndex].IsDown = isDown;
        _fingers[fingerIndex].Id = pointerId;

        // Policz aktywne palce
        _fingerCount = 0;
        for (int i = 0; i < _fingers.Length; i++)
        {
            if (_fingers[i].IsDown)
                _fingerCount++;
        }

        if (isDown && !wasDown)
        {
            // Palec dotknął
            OnFingerDown(fingerIndex, x, y);
        }
        else if (!isDown && wasDown)
        {
            // Palec oderwany
            OnFingerUp(fingerIndex, x, y);
        }
        else if (isDown)
        {
            // Palec się porusza
            OnFingerMove(fingerIndex, x, y);
        }
    }

    private void OnFingerDown(int finger, int x, int y)
    {
        if (_fingerCount == 1)
        {
            // Pierwszy palec - początek gestu
            _gestureStart = new TouchPoint { X = x, Y = y, IsDown = true };
            _touchStartTime = DateTime.Now;
            _gestureExecuted = false;
        }
    }

    private void OnFingerUp(int finger, int x, int y)
    {
        if (_gestureExecuted)
        {
            _gestureExecuted = false;
            return;
        }

        var elapsed = (DateTime.Now - _touchStartTime).TotalMilliseconds;
        int dx = x - _gestureStart.X;
        int dy = y - _gestureStart.Y;
        int absDx = Math.Abs(dx);
        int absDy = Math.Abs(dy);

        // Sprawdź liczbę palców która była aktywna
        int fingers = _fingerCount + 1; // +1 bo ten palec właśnie się oderwał

        // Sprawdź czy to tap czy swipe
        if (absDx < TapThreshold && absDy < TapThreshold && elapsed < SwipeMaxTimeMs)
        {
            // Tap
            DetectTap(fingers, x, y);
        }
        else if (elapsed < SwipeMaxTimeMs)
        {
            // Swipe
            DetectSwipe(fingers, dx, dy, x, y);
        }
    }

    private void OnFingerMove(int finger, int x, int y)
    {
        if (_fingerCount == 1 && !_gestureExecuted)
        {
            // Eksploracja jednym palcem
            ExploreMove?.Invoke(x, y);

            // Sprawdź swipe w trakcie ruchu
            int dx = x - _gestureStart.X;
            int dy = y - _gestureStart.Y;
            int absDx = Math.Abs(dx);
            int absDy = Math.Abs(dy);

            if (absDx > SwipeThreshold || absDy > SwipeThreshold)
            {
                var elapsed = (DateTime.Now - _touchStartTime).TotalMilliseconds;
                if (elapsed < SwipeMaxTimeMs)
                {
                    DetectSwipe(1, dx, dy, x, y);
                    _gestureExecuted = true;
                }
            }
        }
    }

    private void DetectTap(int fingers, int x, int y)
    {
        var now = DateTime.Now;
        bool isDoubleTap = (now - _lastTapTime).TotalMilliseconds < DoubleTapTimeMs;
        _lastTapTime = now;

        TouchpadGesture gesture = TouchpadGesture.None;

        switch (fingers)
        {
            case 1:
                gesture = isDoubleTap ? TouchpadGesture.DoubleTap : TouchpadGesture.SingleTap;
                break;
            case 2:
                gesture = isDoubleTap ? TouchpadGesture.TwoFingerDoubleTap : TouchpadGesture.TwoFingerTap;
                break;
            case 3:
                gesture = TouchpadGesture.ThreeFingerTap;
                break;
        }

        if (gesture != TouchpadGesture.None)
        {
            GestureDetected?.Invoke(gesture, x, y);
        }
    }

    private void DetectSwipe(int fingers, int dx, int dy, int x, int y)
    {
        int absDx = Math.Abs(dx);
        int absDy = Math.Abs(dy);

        bool isHorizontal = absDx > absDy;
        bool isPositive = isHorizontal ? dx > 0 : dy > 0;

        TouchpadGesture gesture = TouchpadGesture.None;

        switch (fingers)
        {
            case 1:
                if (isHorizontal)
                    gesture = isPositive ? TouchpadGesture.SwipeRight : TouchpadGesture.SwipeLeft;
                else
                    gesture = isPositive ? TouchpadGesture.SwipeDown : TouchpadGesture.SwipeUp;
                break;
            case 2:
                if (isHorizontal)
                    gesture = isPositive ? TouchpadGesture.TwoFingerSwipeRight : TouchpadGesture.TwoFingerSwipeLeft;
                else
                    gesture = isPositive ? TouchpadGesture.TwoFingerSwipeDown : TouchpadGesture.TwoFingerSwipeUp;
                break;
            case 3:
                if (isHorizontal)
                    gesture = isPositive ? TouchpadGesture.ThreeFingerSwipeRight : TouchpadGesture.ThreeFingerSwipeLeft;
                break;
        }

        if (gesture != TouchpadGesture.None)
        {
            GestureDetected?.Invoke(gesture, x, y);
        }
    }

    /// <summary>
    /// Symuluje wykrycie gestu (dla testów lub fallback z myszki).
    /// </summary>
    public void SimulateGesture(TouchpadGesture gesture, int x = 0, int y = 0)
    {
        GestureDetected?.Invoke(gesture, x, y);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
    }
}
