using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ScreenReader.VirtualScreen;

/// <summary>
/// Ukryte okno przechwytujące gesty wielopalcowe z touchpada.
/// Używa WM_POINTER API (Windows 8+) dla precyzyjnego wykrywania dotyku.
/// </summary>
public class TouchInputWindow : NativeWindow, IDisposable
{
    private bool _disposed;
    private readonly TouchpadGestureManager _gestureManager;

    // Śledzenie palców
    private readonly Dictionary<int, PointerInfo> _activePointers = new();
    private int _fingerCount;

    // Stan gestu
    private DateTime _gestureStartTime;
    private int _gestureStartFingerCount;
    private Point _gestureStartCenter;
    private bool _gestureInProgress;
    private bool _swipeExecuted;

    // Parametry (jak na telefonach Android/iOS - szybkie gesty)
    private const int SwipeThreshold = 25;      // Mniejsza odległość = szybszy swipe
    private const int TapThreshold = 20;        // Większy margines dla tap
    private const int SwipeMaxTimeMs = 300;     // Więcej czasu = łatwiejszy swipe
    private const int TapMaxTimeMs = 250;       // Więcej czasu dla tap
    private const int DoubleTapTimeMs = 400;    // Więcej czasu dla double-tap

    // Tap tracking
    private DateTime _lastTapTime;
    private int _lastTapFingerCount;
    private Point _lastTapPosition;

    public event Action<TouchpadGesture, int, int>? GestureDetected;
    public event Action<int, int>? ExploreMove;

    #region P/Invoke

    private const int WM_POINTERDOWN = 0x0246;
    private const int WM_POINTERUP = 0x0247;
    private const int WM_POINTERUPDATE = 0x0245;
    private const int WM_POINTERENTER = 0x0249;
    private const int WM_POINTERLEAVE = 0x024A;

    private const int WM_TOUCH = 0x0240;

    [DllImport("user32.dll")]
    private static extern bool GetPointerInfo(uint pointerId, out POINTER_INFO pointerInfo);

    [DllImport("user32.dll")]
    private static extern bool GetPointerTouchInfo(uint pointerId, out POINTER_TOUCH_INFO touchInfo);

    [DllImport("user32.dll")]
    private static extern bool RegisterTouchWindow(IntPtr hwnd, uint ulFlags);

    [DllImport("user32.dll")]
    private static extern bool UnregisterTouchWindow(IntPtr hwnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterPointerInputTarget(IntPtr hwnd, uint pointerType);

    private const uint PT_TOUCH = 0x00000002;
    private const uint PT_TOUCHPAD = 0x00000004;

    private const uint TWF_WANTPALM = 0x00000002;

    [StructLayout(LayoutKind.Sequential)]
    private struct POINTER_INFO
    {
        public uint pointerType;
        public uint pointerId;
        public uint frameId;
        public uint pointerFlags;
        public IntPtr sourceDevice;
        public IntPtr hwndTarget;
        public POINT ptPixelLocation;
        public POINT ptHimetricLocation;
        public POINT ptPixelLocationRaw;
        public POINT ptHimetricLocationRaw;
        public uint dwTime;
        public uint historyCount;
        public int inputData;
        public uint dwKeyStates;
        public ulong PerformanceCount;
        public uint ButtonChangeType;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINTER_TOUCH_INFO
    {
        public POINTER_INFO pointerInfo;
        public uint touchFlags;
        public uint touchMask;
        public RECT rcContact;
        public RECT rcContactRaw;
        public uint orientation;
        public uint pressure;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    private const uint POINTER_FLAG_DOWN = 0x00010000;
    private const uint POINTER_FLAG_UP = 0x00040000;
    private const uint POINTER_FLAG_UPDATE = 0x00000002;
    private const uint POINTER_FLAG_INCONTACT = 0x00000004;

    #endregion

    private struct PointerInfo
    {
        public int Id;
        public int StartX;
        public int StartY;
        public int CurrentX;
        public int CurrentY;
        public DateTime StartTime;
    }

    public TouchInputWindow(TouchpadGestureManager gestureManager)
    {
        _gestureManager = gestureManager;

        // Stwórz ukryte okno
        var cp = new CreateParams
        {
            Caption = "ScreenReader_TouchInput",
            Style = unchecked((int)0x80000000), // WS_POPUP
            ExStyle = 0x08000000, // WS_EX_NOACTIVATE
            X = 0,
            Y = 0,
            Width = 1,
            Height = 1
        };

        CreateHandle(cp);

        // Zarejestruj okno dla dotyku
        if (Handle != IntPtr.Zero)
        {
            RegisterTouchWindow(Handle, TWF_WANTPALM);
            RegisterPointerInputTarget(Handle, PT_TOUCH);
            RegisterPointerInputTarget(Handle, PT_TOUCHPAD);
            Console.WriteLine("TouchInputWindow: Zarejestrowano okno dla WM_POINTER");
        }
    }

    protected override void WndProc(ref Message m)
    {
        switch (m.Msg)
        {
            case WM_POINTERDOWN:
                HandlePointerDown(m.WParam, m.LParam);
                break;

            case WM_POINTERUP:
                HandlePointerUp(m.WParam, m.LParam);
                break;

            case WM_POINTERUPDATE:
                HandlePointerUpdate(m.WParam, m.LParam);
                break;
        }

        base.WndProc(ref m);
    }

    private void HandlePointerDown(IntPtr wParam, IntPtr lParam)
    {
        uint pointerId = (uint)(wParam.ToInt64() & 0xFFFF);

        if (GetPointerInfo(pointerId, out var info))
        {
            if (info.pointerType != PT_TOUCH && info.pointerType != PT_TOUCHPAD)
                return;

            int x = info.ptPixelLocation.X;
            int y = info.ptPixelLocation.Y;

            _activePointers[(int)pointerId] = new PointerInfo
            {
                Id = (int)pointerId,
                StartX = x,
                StartY = y,
                CurrentX = x,
                CurrentY = y,
                StartTime = DateTime.Now
            };

            _fingerCount = _activePointers.Count;

            // Początek nowego gestu
            if (_fingerCount == 1 || !_gestureInProgress)
            {
                _gestureStartTime = DateTime.Now;
                _gestureStartFingerCount = _fingerCount;
                _gestureStartCenter = GetCenterPoint();
                _gestureInProgress = true;
                _swipeExecuted = false;
            }
            else
            {
                // Dodano palec - zaktualizuj
                _gestureStartFingerCount = Math.Max(_gestureStartFingerCount, _fingerCount);
                _gestureStartCenter = GetCenterPoint();
            }

            Console.WriteLine($"PointerDown: id={pointerId}, fingers={_fingerCount}, pos=({x},{y})");
        }
    }

    private void HandlePointerUp(IntPtr wParam, IntPtr lParam)
    {
        uint pointerId = (uint)(wParam.ToInt64() & 0xFFFF);

        if (GetPointerInfo(pointerId, out var info) &&
            info.pointerType != PT_TOUCH &&
            info.pointerType != PT_TOUCHPAD)
        {
            return;
        }

        if (_activePointers.TryGetValue((int)pointerId, out var pointer))
        {
            _activePointers.Remove((int)pointerId);
            _fingerCount = _activePointers.Count;

            Console.WriteLine($"PointerUp: id={pointerId}, remaining={_fingerCount}");

            // Jeśli wszystkie palce oderwane, zakończ gest
            if (_fingerCount == 0 && _gestureInProgress)
            {
                EndGesture(pointer.CurrentX, pointer.CurrentY);
            }
        }
    }

    private void HandlePointerUpdate(IntPtr wParam, IntPtr lParam)
    {
        uint pointerId = (uint)(wParam.ToInt64() & 0xFFFF);

        if (GetPointerInfo(pointerId, out var info))
        {
            if (info.pointerType != PT_TOUCH && info.pointerType != PT_TOUCHPAD)
                return;

            int x = info.ptPixelLocation.X;
            int y = info.ptPixelLocation.Y;

            if (_activePointers.TryGetValue((int)pointerId, out var pointer))
            {
                pointer.CurrentX = x;
                pointer.CurrentY = y;
                _activePointers[(int)pointerId] = pointer;
            }

            if (_gestureInProgress && !_swipeExecuted)
            {
                ProcessGestureMove();
            }

            // Eksploracja 1 palcem
            if (_fingerCount == 1)
            {
                Console.WriteLine($"TouchInputWindow: Eksploracja 1 palcem x={x}, y={y}");
                ExploreMove?.Invoke(x, y);
            }
        }
    }

    private void ProcessGestureMove()
    {
        var center = GetCenterPoint();
        int dx = center.X - _gestureStartCenter.X;
        int dy = center.Y - _gestureStartCenter.Y;
        int absDx = Math.Abs(dx);
        int absDy = Math.Abs(dy);
        var elapsed = (DateTime.Now - _gestureStartTime).TotalMilliseconds;

        // Sprawdź swipe (jak na telefonach - szybki gest)
        if ((absDx > SwipeThreshold || absDy > SwipeThreshold) && elapsed < SwipeMaxTimeMs)
        {
            Console.WriteLine($"TouchInputWindow: SWIPE wykryty! dx={dx}, dy={dy}, elapsed={elapsed}ms, fingers={_gestureStartFingerCount}");
            ExecuteSwipe(_gestureStartFingerCount, dx, dy, absDx, absDy);
            _swipeExecuted = true;
        }
    }

    private void EndGesture(int x, int y)
    {
        _gestureInProgress = false;

        if (_swipeExecuted)
        {
            _swipeExecuted = false;
            return;
        }

        // Sprawdź tap
        var center = GetCenterPoint();
        if (center.X == 0 && center.Y == 0)
        {
            center = new Point(x, y);
        }

        int dx = center.X - _gestureStartCenter.X;
        int dy = center.Y - _gestureStartCenter.Y;
        var elapsed = (DateTime.Now - _gestureStartTime).TotalMilliseconds;

        if (Math.Abs(dx) < TapThreshold && Math.Abs(dy) < TapThreshold && elapsed < TapMaxTimeMs)
        {
            DetectTap(_gestureStartFingerCount, center.X, center.Y);
        }
    }

    private void DetectTap(int fingerCount, int x, int y)
    {
        var now = DateTime.Now;

        // Sprawdź double-tap
        bool isDoubleTap = (now - _lastTapTime).TotalMilliseconds < DoubleTapTimeMs &&
                          fingerCount == _lastTapFingerCount &&
                          Math.Abs(x - _lastTapPosition.X) < TapThreshold * 2 &&
                          Math.Abs(y - _lastTapPosition.Y) < TapThreshold * 2;

        _lastTapTime = now;
        _lastTapFingerCount = fingerCount;
        _lastTapPosition = new Point(x, y);

        TouchpadGesture gesture = TouchpadGesture.None;

        switch (fingerCount)
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
            Console.WriteLine($"Tap detected: {gesture} at ({x},{y})");
            GestureDetected?.Invoke(gesture, x, y);
            _gestureManager.SimulateGesture(gesture, x, y);
        }
    }

    private void ExecuteSwipe(int fingerCount, int dx, int dy, int absDx, int absDy)
    {
        bool isHorizontal = absDx > absDy;
        bool isPositive = isHorizontal ? dx > 0 : dy > 0;

        TouchpadGesture gesture = TouchpadGesture.None;

        switch (fingerCount)
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
            Console.WriteLine($"Swipe detected: {gesture}");
            GestureDetected?.Invoke(gesture, _gestureStartCenter.X, _gestureStartCenter.Y);
            _gestureManager.SimulateGesture(gesture, _gestureStartCenter.X, _gestureStartCenter.Y);
        }
    }

    private Point GetCenterPoint()
    {
        if (_activePointers.Count == 0)
            return new Point(0, 0);

        int sumX = 0, sumY = 0;
        foreach (var p in _activePointers.Values)
        {
            sumX += p.CurrentX;
            sumY += p.CurrentY;
        }

        return new Point(sumX / _activePointers.Count, sumY / _activePointers.Count);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (Handle != IntPtr.Zero)
        {
            UnregisterTouchWindow(Handle);
            DestroyHandle();
        }
    }
}
