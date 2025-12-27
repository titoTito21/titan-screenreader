using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ScreenReader.Keyboard;

/// <summary>
/// Zarządza globalnym hookiem klawiatury
/// Rozszerzona wersja z obsługą flagi LLKHF_EXTENDED dla Insert
/// i nawigacją obiektową przez NumPad (jak NVDA)
/// </summary>
public class KeyboardHookManager : IDisposable
{
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_KEYUP = 0x0101;
    private const int WM_SYSKEYDOWN = 0x0104;
    private const int WM_SYSKEYUP = 0x0105;

    // Flagi z KBDLLHOOKSTRUCT
    private const int LLKHF_EXTENDED = 0x01;
    private const int LLKHF_INJECTED = 0x10;
    private const int LLKHF_ALTDOWN = 0x20;
    private const int LLKHF_UP = 0x80;

    // Virtual key codes
    private const int VK_LEFT = 0x25;
    private const int VK_UP = 0x26;
    private const int VK_RIGHT = 0x27;
    private const int VK_DOWN = 0x28;
    private const int VK_CLEAR = 0x0C;      // NumPad 5 (przy wyłączonym NumLock)
    private const int VK_NUMLOCK = 0x90;
    private const int VK_NUMPAD0 = 0x60;
    private const int VK_NUMPAD1 = 0x61;
    private const int VK_NUMPAD2 = 0x62;
    private const int VK_NUMPAD3 = 0x63;
    private const int VK_NUMPAD4 = 0x64;
    private const int VK_NUMPAD5 = 0x65;
    private const int VK_NUMPAD6 = 0x66;
    private const int VK_NUMPAD7 = 0x67;
    private const int VK_NUMPAD8 = 0x68;
    private const int VK_NUMPAD9 = 0x69;
    private const int VK_RETURN = 0x0D;     // Enter
    private const int VK_ADD = 0x6B;        // NumPad +
    private const int VK_SUBTRACT = 0x6D;   // NumPad -

    // Klawisze nawigacyjne współdzielone z NumPadem
    private const int VK_HOME = 0x24;       // NumPad 7 (przy wyłączonym NumLock)
    private const int VK_END = 0x23;        // NumPad 1 (przy wyłączonym NumLock)
    private const int VK_PRIOR = 0x21;      // Page Up / NumPad 9
    private const int VK_NEXT = 0x22;       // Page Down / NumPad 3

    private LowLevelKeyboardProc? _proc;
    private IntPtr _hookID = IntPtr.Zero;
    private bool _disposed;

    // Stan modyfikatorów
    private bool _ctrlPressed;
    private bool _altPressed;
    private bool _shiftPressed;
    private bool _insertPressed;
    private InsertKeyHandler.InsertKeyType _lastInsertType;

    // Konfiguracja modyfikatora NVDA
    public NVDAModifierConfig NVDAModifierConfig { get; set; } = NVDAModifierConfig.Default;

    /// <summary>
    /// Czy nawigacja numpadem jest włączona (działa przy wyłączonym NumLock)
    /// </summary>
    public bool NumpadNavigationEnabled { get; set; } = true;

    // Eventy dla nawigacji
    public event Action? ReadCurrentElement;
    public event Action? MoveToNextElement;
    public event Action? MoveToPreviousElement;
    public event Action? MoveToParent;
    public event Action? MoveToFirstChild;
    public event Action? StopSpeaking;
    public event Action? ClickAction;
    public event Action? ShowMenu;

    // Eventy dla pól edycyjnych
    public event Action? MoveToPreviousCharacter;
    public event Action? MoveToNextCharacter;
    public event Action? MoveToPreviousLine;
    public event Action? MoveToNextLine;
    public event Action? MoveToPreviousWord;
    public event Action? MoveToNextWord;
    public event Action? MoveToStart;
    public event Action? MoveToEnd;
    public event Action? ReadCurrentChar;
    public event Action? ReadCurrentWord;
    public event Action? ReadCurrentLine;
    public event Action? ReadPosition;

    /// <summary>Event dla wpisanego znaku (echo klawiatury)</summary>
    public event Action<char>? CharTyped;

    /// <summary>Event dla wpisanego słowa (echo klawiatury)</summary>
    public event Action<string>? WordTyped;

    /// <summary>Czy czytnik jest w polu edycyjnym</summary>
    public bool IsInEditField { get; set; }

    /// <summary>Czy czytnik jest w terminalu</summary>
    public bool IsInTerminal { get; set; }

    // Eventy dla terminali - nawigacja tekstowa NumPad
    /// <summary>NumPad 8 - poprzednia linia</summary>
    public event Func<string>? TerminalPreviousLine;
    /// <summary>NumPad 2 - następna linia</summary>
    public event Func<string>? TerminalNextLine;
    /// <summary>NumPad 4 - poprzedni znak</summary>
    public event Func<string>? TerminalPreviousChar;
    /// <summary>NumPad 6 - następny znak</summary>
    public event Func<string>? TerminalNextChar;
    /// <summary>NumPad 1 - poprzedni wyraz</summary>
    public event Func<string>? TerminalPreviousWord;
    /// <summary>NumPad 3 - następny wyraz</summary>
    public event Func<string>? TerminalNextWord;
    /// <summary>NumPad 7 - poprzednia strona</summary>
    public event Func<string>? TerminalPreviousPage;
    /// <summary>NumPad 9 - następna strona</summary>
    public event Func<string>? TerminalNextPage;
    /// <summary>NumPad 5 - odczytaj bieżącą linię</summary>
    public event Func<string>? TerminalReadLine;

    // Bufor dla budowania słowa
    private readonly System.Text.StringBuilder _wordBuffer = new();

    /// <summary>
    /// Event przetwarzania gestów - otrzymuje (vkCode, flags, ctrl, alt, shift, nvdaModifier)
    /// Zwraca true jeśli gest został obsłużony
    /// </summary>
    public event Func<int, int, bool, bool, bool, bool, bool>? GestureProcessed;

    /// <summary>
    /// Event dla szybkiej nawigacji browse mode - otrzymuje (klawisz, shift)
    /// Zwraca true jeśli obsłużony
    /// </summary>
    public event Func<char, bool, bool>? QuickNavProcessed;

    public void Start()
    {
        _proc = HookCallback;
        _hookID = SetHook(_proc);
        Console.WriteLine("KeyboardHookManager: Hook zainstalowany");
    }

    private IntPtr SetHook(LowLevelKeyboardProc proc)
    {
        using var curProcess = Process.GetCurrentProcess();
        using var curModule = curProcess.MainModule;

        if (curModule != null)
        {
            return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                GetModuleHandle(curModule.ModuleName), 0);
        }

        return IntPtr.Zero;
    }

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern short GetKeyState(int nVirtKey);

    /// <summary>
    /// Sprawdza czy NumLock jest wyłączony
    /// </summary>
    private static bool IsNumLockOff()
    {
        return (GetKeyState(VK_NUMLOCK) & 0x0001) == 0;
    }

    /// <summary>
    /// Sprawdza czy klawisz pochodzi z NumPada (nie ma flagi EXTENDED)
    /// Przy wyłączonym NumLock, strzałki z głównej klawiatury mają flagę EXTENDED,
    /// a te z numpada - nie mają
    /// </summary>
    private static bool IsNumpadKey(int vkCode, int flags)
    {
        // Klawisz to nawigacyjny i NIE ma flagi EXTENDED = jest z numpada
        bool isExtended = (flags & LLKHF_EXTENDED) != 0;

        // Te kody są współdzielone między numpadem a strzałkami/nawigacją
        // NumPad 2/4/6/8 = strzałki, NumPad 5 = Clear
        // NumPad 1/3/7/9 = End/Next/Home/Prior
        bool isNavigationKey = vkCode == VK_LEFT || vkCode == VK_RIGHT ||
                               vkCode == VK_UP || vkCode == VK_DOWN ||
                               vkCode == VK_CLEAR ||
                               vkCode == VK_HOME || vkCode == VK_END ||
                               vkCode == VK_PRIOR || vkCode == VK_NEXT;

        // NumPad Enter też nie ma flagi EXTENDED (główny Enter ma)
        bool isNumpadEnter = vkCode == VK_RETURN && !isExtended;

        return (isNavigationKey && !isExtended) || isNumpadEnter;
    }

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        try
        {
            if (nCode >= 0)
            {
                var hookStruct = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                int vkCode = (int)hookStruct.vkCode;
                int flags = (int)hookStruct.flags;

                bool isKeyDown = wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN;
                bool isKeyUp = wParam == (IntPtr)WM_KEYUP || wParam == (IntPtr)WM_SYSKEYUP;

                // Aktualizuj stan modyfikatorów
                UpdateModifierState(vkCode, flags, isKeyDown);

                if (isKeyDown)
                {
                    bool handled = ProcessKeyDown(vkCode, flags);
                    if (handled)
                        return (IntPtr)1; // Blokuj klawisz
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"KeyboardHookManager: Błąd w hook: {ex.Message}");
        }

        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    /// <summary>
    /// Aktualizuje stan modyfikatorów
    /// </summary>
    private void UpdateModifierState(int vkCode, int flags, bool isDown)
    {
        switch (vkCode)
        {
            case 0x11: // VK_CONTROL
            case 0xA2: // VK_LCONTROL
            case 0xA3: // VK_RCONTROL
                _ctrlPressed = isDown;
                break;

            case 0x12: // VK_MENU (Alt)
            case 0xA4: // VK_LMENU
            case 0xA5: // VK_RMENU
                _altPressed = isDown;
                break;

            case 0x10: // VK_SHIFT
            case 0xA0: // VK_LSHIFT
            case 0xA1: // VK_RSHIFT
                _shiftPressed = isDown;
                break;

            case InsertKeyHandler.VK_INSERT:
                if (isDown)
                {
                    _lastInsertType = InsertKeyHandler.GetInsertKeyType(vkCode, flags);
                    _insertPressed = InsertKeyHandler.IsNVDAModifierKey(vkCode, flags, NVDAModifierConfig);
                }
                else
                {
                    _insertPressed = false;
                }
                break;

            case InsertKeyHandler.VK_CAPSLOCK:
                if (NVDAModifierConfig.HasFlag(NVDAModifierConfig.CapsLock))
                {
                    _insertPressed = isDown;
                }
                break;
        }
    }

    /// <summary>
    /// Przetwarza naciśnięcie klawisza
    /// </summary>
    private bool ProcessKeyDown(int vkCode, int flags)
    {
        // Sprawdź czy to klawisz Insert sam w sobie
        if (vkCode == InsertKeyHandler.VK_INSERT)
        {
            // Nie blokuj samego Insert
            return false;
        }

        // Sprawdź gesty (Insert+...)
        if (_insertPressed && GestureProcessed != null)
        {
            bool handled = GestureProcessed(vkCode, flags, _ctrlPressed, _altPressed, _shiftPressed, _insertPressed);
            if (handled)
                return true;
        }

        // Sprawdź szybką nawigację (pojedyncze litery w browse mode)
        if (!_ctrlPressed && !_altPressed && !_insertPressed)
        {
            char? ch = VkCodeToChar(vkCode);
            if (ch.HasValue && QuickNavProcessed != null)
            {
                bool handled = QuickNavProcessed(ch.Value, _shiftPressed);
                if (handled)
                    return true;
            }
        }

        // Ctrl+Shift+Backslash - menu
        if (_ctrlPressed && _shiftPressed && vkCode == 0xDC)
        {
            ShowMenu?.Invoke();
            return true;
        }

        // Enter - akcja kliknięcia (nie blokuj)
        if (vkCode == 0x0D)
        {
            ClickAction?.Invoke();
            return false;
        }

        // ========================================
        // NAWIGACJA NUMPADEM (jak w NVDA)
        // Działa przy wyłączonym NumLock
        // NumPad 4 = lewo (poprzedni element)
        // NumPad 6 = prawo (następny element)
        // NumPad 8 = góra (rodzic)
        // NumPad 2 = dół (potomek)
        // NumPad 5 = odczytaj bieżący element
        // ========================================
        if (NumpadNavigationEnabled && IsNumLockOff() && IsNumpadKey(vkCode, flags))
        {
            bool handled = ProcessNumpadNavigation(vkCode);
            if (handled)
                return true;
        }

        // Ctrl+Alt+... - skróty czytnika ekranu (blokują klawisze)
        // Zachowane dla kompatybilności wstecznej
        if (_ctrlPressed && _altPressed)
        {
            return ProcessCtrlAltShortcut(vkCode);
        }

        // Ctrl+strzałki (bez Alt) - nawigacja po słowach
        // NIE blokujemy - pozwalamy systemowi przesunąć kursor, ale ogłaszamy słowo
        if (_ctrlPressed && !_altPressed && !_insertPressed)
        {
            ProcessCtrlArrowNavigation(vkCode);
            // Zwracamy false - nie blokujemy, klawisz idzie do systemu
        }

        // Strzałki bez modyfikatorów w polu edycyjnym - nawigacja po znakach/liniach
        // NIE blokujemy - pozwalamy systemowi przesunąć kursor, ale ogłaszamy znak/linię
        if (!_ctrlPressed && !_altPressed && !_insertPressed && IsInEditField)
        {
            ProcessArrowNavigation(vkCode);
            // Zwracamy false - nie blokujemy, klawisz idzie do systemu
        }

        // Echo klawiatury - przechwytuj wpisywane znaki (tylko w polu edycyjnym)
        if (IsInEditField && !_ctrlPressed && !_altPressed && !_insertPressed)
        {
            ProcessTypedCharacter(vkCode);
        }

        return false;
    }

    /// <summary>
    /// Przetwarza wpisany znak dla echa klawiatury
    /// </summary>
    private void ProcessTypedCharacter(int vkCode)
    {
        // Spacja, Enter, Tab - zakończ słowo
        if (vkCode == 0x20 || vkCode == 0x0D || vkCode == 0x09)
        {
            FlushWordBuffer();

            // Ogłoś spację jako "spacja"
            if (vkCode == 0x20)
            {
                CharTyped?.Invoke(' ');
            }
            else if (vkCode == 0x0D)
            {
                CharTyped?.Invoke('\n');
            }
            return;
        }

        // Backspace - usuń ostatni znak z bufora
        if (vkCode == 0x08)
        {
            if (_wordBuffer.Length > 0)
            {
                _wordBuffer.Remove(_wordBuffer.Length - 1, 1);
            }
            CharTyped?.Invoke('\b');
            return;
        }

        // Interpunkcja - zakończ słowo i ogłoś znak
        if (IsPunctuation(vkCode))
        {
            FlushWordBuffer();
            var ch = VkCodeToTypedChar(vkCode);
            if (ch.HasValue)
            {
                CharTyped?.Invoke(ch.Value);
            }
            return;
        }

        // Litery i cyfry - dodaj do bufora i ogłoś
        var typedChar = VkCodeToTypedChar(vkCode);
        if (typedChar.HasValue)
        {
            char ch = typedChar.Value;

            // Uwzględnij Shift dla wielkich liter
            if (_shiftPressed && char.IsLetter(ch))
            {
                ch = char.ToUpper(ch);
            }
            else if (!_shiftPressed && char.IsLetter(ch))
            {
                ch = char.ToLower(ch);
            }

            _wordBuffer.Append(ch);
            CharTyped?.Invoke(ch);
        }
    }

    /// <summary>
    /// Wysyła buforowane słowo i czyści bufor
    /// </summary>
    private void FlushWordBuffer()
    {
        if (_wordBuffer.Length > 0)
        {
            WordTyped?.Invoke(_wordBuffer.ToString());
            _wordBuffer.Clear();
        }
    }

    /// <summary>
    /// Sprawdza czy klawisz to interpunkcja
    /// </summary>
    private static bool IsPunctuation(int vkCode)
    {
        // Przecinek, kropka, średnik, dwukropek, itd.
        return vkCode == 0xBC || // , (przecinek)
               vkCode == 0xBE || // . (kropka)
               vkCode == 0xBA || // ; (średnik)
               vkCode == 0xBB || // = (znak równości)
               vkCode == 0xBD || // - (minus)
               vkCode == 0xC0 || // ` (grawis)
               vkCode == 0xDB || // [ (nawias kwadratowy)
               vkCode == 0xDD || // ] (nawias kwadratowy)
               vkCode == 0xDC || // \ (backslash)
               vkCode == 0xDE || // ' (apostrof)
               vkCode == 0xBF;   // / (slash)
    }

    /// <summary>
    /// Konwertuje kod klawisza na wpisany znak
    /// </summary>
    private static char? VkCodeToTypedChar(int vkCode)
    {
        // Litery A-Z
        if (vkCode >= 0x41 && vkCode <= 0x5A)
            return (char)vkCode;

        // Cyfry 0-9
        if (vkCode >= 0x30 && vkCode <= 0x39)
            return (char)vkCode;

        // NumPad cyfry
        if (vkCode >= 0x60 && vkCode <= 0x69)
            return (char)('0' + (vkCode - 0x60));

        // Interpunkcja
        return vkCode switch
        {
            0xBC => ',',
            0xBE => '.',
            0xBA => ';',
            0xBB => '=',
            0xBD => '-',
            0xC0 => '`',
            0xDB => '[',
            0xDD => ']',
            0xDC => '\\',
            0xDE => '\'',
            0xBF => '/',
            _ => null
        };
    }

    /// <summary>
    /// Przetwarza nawigację numpadem (przy wyłączonym NumLock)
    /// </summary>
    private bool ProcessNumpadNavigation(int vkCode)
    {
        // W terminalu używamy nawigacji tekstowej
        if (IsInTerminal)
        {
            return ProcessTerminalNumpadNavigation(vkCode);
        }

        // Standardowa nawigacja obiektowa
        switch (vkCode)
        {
            case VK_LEFT:   // NumPad 4 - poprzedni element (rodzeństwo)
                MoveToPreviousElement?.Invoke();
                return true;

            case VK_RIGHT:  // NumPad 6 - następny element (rodzeństwo)
                MoveToNextElement?.Invoke();
                return true;

            case VK_UP:     // NumPad 8 - element nadrzędny (rodzic)
                MoveToParent?.Invoke();
                return true;

            case VK_DOWN:   // NumPad 2 - pierwszy potomek
                MoveToFirstChild?.Invoke();
                return true;

            case VK_CLEAR:  // NumPad 5 - odczytaj bieżący element
                ReadCurrentElement?.Invoke();
                return true;

            case VK_RETURN: // NumPad Enter - aktywuj element
                ClickAction?.Invoke();
                return true;
        }

        return false;
    }

    /// <summary>
    /// Przetwarza nawigację numpadem w trybie terminala
    /// NumPad 2: następna linia, 8: poprzednia linia
    /// NumPad 4: poprzedni znak, 6: następny znak
    /// NumPad 1: poprzedni wyraz, 3: następny wyraz
    /// NumPad 7: poprzednia strona, 9: następna strona
    /// NumPad 5: odczytaj bieżącą linię
    /// </summary>
    private bool ProcessTerminalNumpadNavigation(int vkCode)
    {
        string? result = null;

        switch (vkCode)
        {
            case VK_DOWN:   // NumPad 2 - następna linia
                result = TerminalNextLine?.Invoke();
                break;

            case VK_UP:     // NumPad 8 - poprzednia linia
                result = TerminalPreviousLine?.Invoke();
                break;

            case VK_LEFT:   // NumPad 4 - poprzedni znak
                result = TerminalPreviousChar?.Invoke();
                break;

            case VK_RIGHT:  // NumPad 6 - następny znak
                result = TerminalNextChar?.Invoke();
                break;

            case VK_END:    // NumPad 1 - poprzedni wyraz
                result = TerminalPreviousWord?.Invoke();
                break;

            case VK_NEXT:   // NumPad 3 - następny wyraz
                result = TerminalNextWord?.Invoke();
                break;

            case VK_HOME:   // NumPad 7 - poprzednia strona
                result = TerminalPreviousPage?.Invoke();
                break;

            case VK_PRIOR:  // NumPad 9 - następna strona
                result = TerminalNextPage?.Invoke();
                break;

            case VK_CLEAR:  // NumPad 5 - odczytaj bieżącą linię
                result = TerminalReadLine?.Invoke();
                break;

            case VK_RETURN: // NumPad Enter - aktywuj element
                ClickAction?.Invoke();
                return true;

            default:
                return false;
        }

        // Wynik będzie obsłużony przez ScreenReaderEngine (mowa)
        return result != null;
    }

    /// <summary>
    /// Przetwarza strzałki dla nawigacji po znakach/liniach (nie blokuje klawiszy)
    /// </summary>
    private void ProcessArrowNavigation(int vkCode)
    {
        switch (vkCode)
        {
            case 0x27: // Right - następny znak
                MoveToNextCharacter?.Invoke();
                break;

            case 0x25: // Left - poprzedni znak
                MoveToPreviousCharacter?.Invoke();
                break;

            case 0x26: // Up - poprzednia linia
                MoveToPreviousLine?.Invoke();
                break;

            case 0x28: // Down - następna linia
                MoveToNextLine?.Invoke();
                break;

            case 0x24: // Home - początek linii
                MoveToStart?.Invoke();
                break;

            case 0x23: // End - koniec linii
                MoveToEnd?.Invoke();
                break;
        }
    }

    /// <summary>
    /// Przetwarza Ctrl+strzałki dla nawigacji po słowach (nie blokuje klawiszy)
    /// </summary>
    private void ProcessCtrlArrowNavigation(int vkCode)
    {
        switch (vkCode)
        {
            case 0x27: // Right - następne słowo
                MoveToNextWord?.Invoke();
                break;

            case 0x25: // Left - poprzednie słowo
                MoveToPreviousWord?.Invoke();
                break;

            case 0x24: // Home - początek dokumentu/linii
                MoveToStart?.Invoke();
                break;

            case 0x23: // End - koniec dokumentu/linii
                MoveToEnd?.Invoke();
                break;
        }
    }

    /// <summary>
    /// Przetwarza skróty Ctrl+Alt+...
    /// </summary>
    private bool ProcessCtrlAltShortcut(int vkCode)
    {
        switch (vkCode)
        {
            case 0x20: // Space - odczytaj bieżący element
                ReadCurrentElement?.Invoke();
                return true;

            case 0x27: // Right - następny element / następny znak w polu edycji
                MoveToNextElement?.Invoke();
                return true;

            case 0x25: // Left - poprzedni element / poprzedni znak w polu edycji
                MoveToPreviousElement?.Invoke();
                return true;

            case 0x26: // Up - element nadrzędny
                MoveToParent?.Invoke();
                return true;

            case 0x28: // Down - pierwszy potomek
                MoveToFirstChild?.Invoke();
                return true;

            case 0x24: // Home - początek
                MoveToStart?.Invoke();
                return true;

            case 0x23: // End - koniec
                MoveToEnd?.Invoke();
                return true;

            case 0x53: // S - zatrzymaj mowę
                StopSpeaking?.Invoke();
                return true;

            case 0x43: // C - odczytaj znak
                ReadCurrentChar?.Invoke();
                return true;

            case 0x57: // W - odczytaj słowo
                ReadCurrentWord?.Invoke();
                return true;

            case 0x4C: // L - odczytaj linię
                ReadCurrentLine?.Invoke();
                return true;

            case 0x50: // P - odczytaj pozycję
                ReadPosition?.Invoke();
                return true;
        }

        return false;
    }

    /// <summary>
    /// Konwertuje kod klawisza na znak (dla szybkiej nawigacji)
    /// </summary>
    private static char? VkCodeToChar(int vkCode)
    {
        // Litery A-Z
        if (vkCode >= 0x41 && vkCode <= 0x5A)
            return (char)vkCode;

        // Cyfry 0-9
        if (vkCode >= 0x30 && vkCode <= 0x39)
            return (char)vkCode;

        return null;
    }

    public void Stop()
    {
        if (_hookID != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hookID);
            _hookID = IntPtr.Zero;
            Console.WriteLine("KeyboardHookManager: Hook usunięty");
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        Stop();
        _disposed = true;
    }

    #region Struktury Windows

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public uint vkCode;
        public uint scanCode;
        public uint flags;      // Zawiera LLKHF_EXTENDED, LLKHF_INJECTED, itp.
        public uint time;
        public IntPtr dwExtraInfo;
    }

    #endregion

    #region Windows API

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook,
        LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
        IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    #endregion
}
