using System.Windows.Automation;
using ScreenReader.InputGestures;
using ScreenReader.AppModules;
using ScreenReader.BrowseMode;
using ScreenReader.EditableText;
using ScreenReader.Keyboard;
using KeyboardHook = ScreenReader.Keyboard.KeyboardHookManager;

namespace ScreenReader;

public class ScreenReaderEngine : IDisposable
{
    // Singleton dla dostępu z innych komponentów
    public static ScreenReaderEngine? Instance { get; private set; }

    private readonly SpeechManager _speechManager;
    private readonly SoundManager _soundManager;
    private readonly FocusTracker _focusTracker;
    private readonly KeyboardHook _keyboardHook;
    private readonly EditFieldNavigator _editNavigator;
    private readonly GestureManager _gestureManager;

    // Nowe komponenty
    private readonly AppModuleManager _appModuleManager;
    private readonly BrowseModeHandler _browseModeHandler;
    private readonly EditableTextHandler _editableTextHandler;

    private DialogMonitor? _dialogMonitor;
    private AutomationElement? _currentElement;
    private AutomationElement? _lastWindow;
    private bool _disposed;

    // Echo klawiatury
    private KeyboardEchoMode _keyboardEchoMode = KeyboardEchoMode.Characters;

    // Menu kontekstowe
    private ScreenReaderContextMenu? _contextMenu;

    /// <summary>Aktualny tryb echa klawiatury</summary>
    public KeyboardEchoMode KeyboardEchoMode => _keyboardEchoMode;

    public ScreenReaderEngine()
    {
        Instance = this;

        _speechManager = new SpeechManager();
        var soundsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sfx");
        _soundManager = new SoundManager(soundsPath);
        _focusTracker = new FocusTracker();
        _keyboardHook = new KeyboardHook();
        _editNavigator = new EditFieldNavigator(_speechManager);
        _gestureManager = new GestureManager(_speechManager);

        // Nowe komponenty
        _appModuleManager = new AppModuleManager();
        _browseModeHandler = new BrowseModeHandler();
        _editableTextHandler = new EditableTextHandler();

        // Podłącz eventy browse mode
        _browseModeHandler.Announce += text => _speechManager.Speak(text);
        _browseModeHandler.ModeChanged += OnBrowseModeChanged;

        // Podłącz eventy edytowalnego tekstu
        _editableTextHandler.Announce += text => _speechManager.Speak(text);

        // Wire up events
        _focusTracker.FocusChanged += OnFocusChanged;
        _keyboardHook.ReadCurrentElement += OnReadCurrentElement;
        _keyboardHook.MoveToNextElement += OnMoveToNextElement;
        _keyboardHook.MoveToPreviousElement += OnMoveToPreviousElement;
        _keyboardHook.MoveToParent += OnMoveToParent;
        _keyboardHook.MoveToFirstChild += OnMoveToFirstChild;
        _keyboardHook.StopSpeaking += OnStopSpeaking;
        _keyboardHook.ClickAction += OnClickAction;
        _keyboardHook.ShowMenu += OnShowMenu;

        // Edit field navigation events
        _keyboardHook.MoveToPreviousCharacter += OnMoveToPreviousCharacter;
        _keyboardHook.MoveToNextCharacter += OnMoveToNextCharacter;
        _keyboardHook.MoveToPreviousLine += OnMoveToPreviousLine;
        _keyboardHook.MoveToNextLine += OnMoveToNextLine;
        _keyboardHook.MoveToPreviousWord += OnMoveToPreviousWord;
        _keyboardHook.MoveToNextWord += OnMoveToNextWord;
        _keyboardHook.MoveToStart += OnMoveToStart;
        _keyboardHook.MoveToEnd += OnMoveToEnd;
        _keyboardHook.ReadCurrentChar += OnReadCurrentChar;
        _keyboardHook.ReadCurrentWord += OnReadCurrentWord;
        _keyboardHook.ReadCurrentLine += OnReadCurrentLine;
        _keyboardHook.ReadPosition += OnReadPosition;

        // Connect gesture processing (nowa sygnatura)
        _keyboardHook.GestureProcessed += OnGestureProcessed;

        // Connect quick nav for browse mode
        _keyboardHook.QuickNavProcessed += OnQuickNavProcessed;

        // Connect keyboard echo events
        _keyboardHook.CharTyped += OnCharTyped;
        _keyboardHook.WordTyped += OnWordTyped;

        // Connect terminal navigation events
        _keyboardHook.TerminalPreviousLine += OnTerminalPreviousLine;
        _keyboardHook.TerminalNextLine += OnTerminalNextLine;
        _keyboardHook.TerminalPreviousChar += OnTerminalPreviousChar;
        _keyboardHook.TerminalNextChar += OnTerminalNextChar;
        _keyboardHook.TerminalPreviousWord += OnTerminalPreviousWord;
        _keyboardHook.TerminalNextWord += OnTerminalNextWord;
        _keyboardHook.TerminalPreviousPage += OnTerminalPreviousPage;
        _keyboardHook.TerminalNextPage += OnTerminalNextPage;
        _keyboardHook.TerminalReadLine += OnTerminalReadLine;

        // Register additional gestures
        RegisterCustomGestures();
    }

    /// <summary>
    /// Obsługa wpisanego znaku (echo klawiatury)
    /// </summary>
    private void OnCharTyped(char ch)
    {
        if (!_keyboardEchoMode.IncludesCharacters())
            return;

        string announcement = GetCharacterAnnouncement(ch);
        if (!string.IsNullOrEmpty(announcement))
        {
            _speechManager.Speak(announcement);
        }
    }

    /// <summary>
    /// Obsługa wpisanego słowa (echo klawiatury)
    /// </summary>
    private void OnWordTyped(string word)
    {
        if (!_keyboardEchoMode.IncludesWords())
            return;

        if (!string.IsNullOrWhiteSpace(word))
        {
            _speechManager.Speak(word);
        }
    }

    /// <summary>
    /// Zwraca ogłoszenie dla znaku (z polskim alfabetem fonetycznym)
    /// </summary>
    private static string GetCharacterAnnouncement(char ch)
    {
        // Znaki specjalne
        return ch switch
        {
            ' ' => "spacja",
            '\n' => "nowa linia",
            '\r' => "",
            '\t' => "tabulator",
            '\b' => "", // Backspace - nie ogłaszaj
            '.' => "kropka",
            ',' => "przecinek",
            ';' => "średnik",
            ':' => "dwukropek",
            '!' => "wykrzyknik",
            '?' => "pytajnik",
            '-' => "minus",
            '_' => "podkreślenie",
            '=' => "równa się",
            '+' => "plus",
            '*' => "gwiazdka",
            '/' => "ukośnik",
            '\\' => "odwrotny ukośnik",
            '@' => "małpa",
            '#' => "hash",
            '$' => "dolar",
            '%' => "procent",
            '^' => "daszek",
            '&' => "ampersand",
            '(' => "nawias otwierający",
            ')' => "nawias zamykający",
            '[' => "nawias kwadratowy otwierający",
            ']' => "nawias kwadratowy zamykający",
            '{' => "nawias klamrowy otwierający",
            '}' => "nawias klamrowy zamykający",
            '<' => "mniejszy niż",
            '>' => "większy niż",
            '\'' => "apostrof",
            '"' => "cudzysłów",
            '`' => "grawis",
            '~' => "tylda",
            '|' => "kreska pionowa",
            // Wielkie litery - ogłoś jako "duże X"
            >= 'A' and <= 'Z' => $"duże {ch}",
            // Małe litery i cyfry - ogłoś sam znak
            _ => ch.ToString()
        };
    }

    /// <summary>
    /// Obsługa gestów (Insert+...)
    /// </summary>
    private bool OnGestureProcessed(int vkCode, int flags, bool ctrl, bool alt, bool shift, bool nvdaModifier)
    {
        if (!nvdaModifier)
            return false;

        // Konwertuj na Keys i przekaż do GestureManager
        var key = (System.Windows.Forms.Keys)vkCode;
        return _gestureManager.ProcessKeyPress(key, ctrl, alt, shift, nvdaModifier);
    }

    /// <summary>
    /// Obsługa szybkiej nawigacji (browse mode)
    /// </summary>
    private bool OnQuickNavProcessed(char key, bool shift)
    {
        if (_browseModeHandler.IsActive)
        {
            return _browseModeHandler.HandleQuickNav(key, shift);
        }
        return false;
    }

    /// <summary>
    /// Obsługa zmiany trybu browse/focus
    /// </summary>
    private void OnBrowseModeChanged(bool passThrough)
    {
        // Odtwórz dźwięk zmiany trybu
        if (passThrough)
        {
            _soundManager.PlayClicked(); // Focus mode
        }
        else
        {
            _soundManager.PlayCursor(); // Browse mode
        }
    }

    /// <summary>
    /// Aktywuje browse mode dla dokumentu
    /// </summary>
    public void ActivateBrowseMode(AutomationElement document)
    {
        _browseModeHandler.Activate(document);
    }

    public void Start()
    {
        Console.WriteLine("Uruchamianie Czytnika Ekranu...");
        _speechManager.Speak("Czytnik ekranu uruchomiony");
        _soundManager.PlayWindow();
        
        _focusTracker.Start();
        _keyboardHook.Start();
        
        // Start dialog monitor
        _dialogMonitor = new DialogMonitor(_speechManager);
        _dialogMonitor.StartMonitoring();

        // Read the currently focused element on startup
        _currentElement = UIAutomationHelper.GetFocusedElement();
        if (_currentElement != null)
        {
            AnnounceElement(_currentElement, false);
        }

        Console.WriteLine("Czytnik Ekranu działa. Naciśnij Ctrl+C aby zakończyć.");
        Console.WriteLine("Skróty klawiszowe (NumPad przy wyłączonym NumLock):");
        Console.WriteLine("  NumPad 5: Odczytaj bieżący element");
        Console.WriteLine("  NumPad 6: Następny element (w prawo)");
        Console.WriteLine("  NumPad 4: Poprzedni element (w lewo)");
        Console.WriteLine("  NumPad 2: Pierwszy potomek (w dół)");
        Console.WriteLine("  NumPad 8: Element nadrzędny (w górę)");
        Console.WriteLine("  NumPad Enter: Aktywuj element");
        Console.WriteLine("  Ctrl+Alt+S: Zatrzymaj mowę");
        Console.WriteLine("  Ctrl+Shift+\\: Menu czytnika");
    }

    private void OnFocusChanged(AutomationElement element)
    {
        try
        {
            // Sprawdź czy element jest dostępny
            try
            {
                _ = element.Current.ProcessId;
            }
            catch (System.Windows.Automation.ElementNotAvailableException)
            {
                return; // Element zniknął
            }

            _currentElement = element;

            // Powiadom AppModuleManager o zmianie fokusu
            _appModuleManager.OnFocusChanged(element);

            // Sprawdź czy moduł aplikacji chce użyć wirtualnego bufora
            var currentModule = _appModuleManager.CurrentModule;
            if (currentModule != null && currentModule.ShouldUseVirtualBuffer(element))
            {
                ActivateBrowseMode(element);
            }
            else
            {
                _browseModeHandler.Deactivate();
            }

            // Check if we're in a terminal - update keyboard hook
            _keyboardHook.IsInTerminal = _appModuleManager.IsTerminalActive;

            // Check if it's an edit field - użyj nowego EditableTextHandler
            bool isEditField = EditableTextHandler.IsEditField(element);
            _keyboardHook.IsInEditField = isEditField;

        if (isEditField)
        {
            _editNavigator.SetCurrentEdit(element);
            _editableTextHandler.SetElement(element);
        }
        else
        {
            _editNavigator.SetCurrentEdit(null);
            _editableTextHandler.SetElement(null);
        }

        // Check if window changed
        var window = GetContainingWindow(element);
        if (window != null && _lastWindow != null)
        {
            try
            {
                if (!Automation.Compare(window, _lastWindow))
                {
                    _lastWindow = window;
                    _soundManager.PlayWindow();
                    var title = GetWindowTitle(window);
                    Console.WriteLine($"Nowe okno: {title}");
                    _speechManager.Speak($"Okno {title}");

                    // Powiadom moduł aplikacji
                    _appModuleManager.OnWindowChanged(window);
                }
            }
            catch
            {
                // Ignore comparison errors
            }
        }
        else if (window != null && _lastWindow == null)
        {
            _lastWindow = window;
        }

            AnnounceElement(element, false);
        }
        catch (System.Windows.Automation.ElementNotAvailableException)
        {
            // Element zniknął podczas przetwarzania - normalne zachowanie
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd w OnFocusChanged: {ex.Message}");
        }
    }

    private void OnReadCurrentElement()
    {
        if (_currentElement != null)
        {
            // Play clicked sound for buttons or when Enter is pressed
            if (UIAutomationHelper.IsButton(_currentElement))
            {
                _soundManager.PlayClicked();
            }
            
            AnnounceElement(_currentElement, false);
        }
        else
        {
            _speechManager.Speak("Brak bieżącego elementu");
            Console.WriteLine("Brak bieżącego elementu");
        }
    }

    private void OnMoveToNextElement()
    {
        var next = UIAutomationHelper.GetNextSibling(_currentElement);
        if (next != null)
        {
            _currentElement = next;
            _soundManager.PlayCursor();
            AnnounceElement(next, true);
        }
        else
        {
            _soundManager.PlayEdge();
            _speechManager.Speak("Koniec listy");
            Console.WriteLine("Koniec listy");
        }
    }

    private void OnMoveToPreviousElement()
    {
        var previous = UIAutomationHelper.GetPreviousSibling(_currentElement);
        if (previous != null)
        {
            _currentElement = previous;
            _soundManager.PlayCursor();
            AnnounceElement(previous, true);
        }
        else
        {
            _soundManager.PlayEdge();
            _speechManager.Speak("Początek listy");
            Console.WriteLine("Początek listy");
        }
    }

    private void OnMoveToParent()
    {
        var parent = UIAutomationHelper.GetParent(_currentElement);
        if (parent != null)
        {
            _currentElement = parent;
            _soundManager.PlayCursor();
            AnnounceElement(parent, false);
        }
        else
        {
            _soundManager.PlayEdge();
            _speechManager.Speak("Brak elementu nadrzędnego");
            Console.WriteLine("Brak elementu nadrzędnego");
        }
    }

    private void OnMoveToFirstChild()
    {
        var child = UIAutomationHelper.GetFirstChild(_currentElement);
        if (child != null)
        {
            _currentElement = child;
            _soundManager.PlayCursor();
            AnnounceElement(child, false);
        }
        else
        {
            _soundManager.PlayEdge();
            _speechManager.Speak("Brak elementów potomnych");
            Console.WriteLine("Brak elementów potomnych");
        }
    }

    private void OnClickAction()
    {
        if (_currentElement == null)
        {
            _speechManager.Speak("Brak bieżącego elementu");
            return;
        }

        try
        {
            // Spróbuj aktywować element przez InvokePattern
            if (_currentElement.TryGetCurrentPattern(
                System.Windows.Automation.InvokePattern.Pattern, out var invokePattern))
            {
                _soundManager.PlayClicked();
                ((System.Windows.Automation.InvokePattern)invokePattern).Invoke();
                Console.WriteLine("Akcja: Aktywowano element");
                return;
            }

            // Dla pól wyboru użyj TogglePattern
            if (_currentElement.TryGetCurrentPattern(
                System.Windows.Automation.TogglePattern.Pattern, out var togglePattern))
            {
                _soundManager.PlayClicked();
                var toggle = (System.Windows.Automation.TogglePattern)togglePattern;
                toggle.Toggle();
                var state = toggle.Current.ToggleState == System.Windows.Automation.ToggleState.On
                    ? "zaznaczone"
                    : "odznaczone";
                _speechManager.Speak(state);
                Console.WriteLine($"Akcja: Toggle - {state}");
                return;
            }

            // Dla elementów wyboru użyj SelectionItemPattern
            if (_currentElement.TryGetCurrentPattern(
                System.Windows.Automation.SelectionItemPattern.Pattern, out var selectionPattern))
            {
                _soundManager.PlayClicked();
                ((System.Windows.Automation.SelectionItemPattern)selectionPattern).Select();
                Console.WriteLine("Akcja: Wybrano element");
                return;
            }

            // Dla elementów rozwijalnych użyj ExpandCollapsePattern
            if (_currentElement.TryGetCurrentPattern(
                System.Windows.Automation.ExpandCollapsePattern.Pattern, out var expandPattern))
            {
                _soundManager.PlayClicked();
                var expand = (System.Windows.Automation.ExpandCollapsePattern)expandPattern;
                if (expand.Current.ExpandCollapseState == System.Windows.Automation.ExpandCollapseState.Expanded)
                {
                    expand.Collapse();
                    _speechManager.Speak("zwinięty");
                }
                else
                {
                    expand.Expand();
                    _speechManager.Speak("rozwinięty");
                }
                Console.WriteLine("Akcja: Rozwiń/zwiń");
                return;
            }

            // Fallback: spróbuj ustawić fokus
            _currentElement.SetFocus();
            _soundManager.PlayClicked();
            Console.WriteLine("Akcja: Ustawiono fokus");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd aktywacji elementu: {ex.Message}");
            _speechManager.Speak("Nie można aktywować elementu");
        }
    }

    private void OnShowMenu()
    {
        Console.WriteLine("Menu czytnika ekranu");
        _speechManager.Speak("Menu");

        // Show context menu on UI thread
        if (System.Windows.Forms.Application.MessageLoop)
        {
            _contextMenu ??= new ScreenReaderContextMenu(
                onSettings: ShowSettings,
                onHelp: ShowHelp,
                onExit: () =>
                {
                    Console.WriteLine("Zamykanie czytnika ekranu...");
                    _speechManager.Speak("Zamykanie czytnika ekranu");
                    System.Windows.Forms.Application.Exit();
                }
            );

            _contextMenu.ShowCentered();
        }
    }

    private void ShowHelp()
    {
        Console.WriteLine("Otwieranie pomocy...");
        _speechManager.Speak("Pomoc czytnika ekranu. Naciśnij Insert+1 aby włączyć pomoc klawiatury.");
    }

    private void ShowSettings()
    {
        Console.WriteLine("Otwieranie ustawień...");
        
        if (System.Windows.Forms.Application.MessageLoop)
        {
            var settingsDialog = new SettingsDialog(_speechManager);
            settingsDialog.ShowDialog();
        }
    }

    private void OnStopSpeaking()
    {
        _speechManager.Stop();
        _soundManager.Stop();
        Console.WriteLine("Dźwięk zatrzymany");
    }

    // Edit field navigation handlers
    // Używamy Task.Run z opóźnieniem, bo event jest wywoływany PRZED przesunięciem kursora
    private void OnMoveToPreviousCharacter()
    {
        // Strzałka w lewo - ogłoś znak na który przeszliśmy
        Task.Run(async () =>
        {
            await Task.Delay(50); // Poczekaj aż kursor się przesunie
            _editableTextHandler.ReadCurrentCharacter();
        });
    }

    private void OnMoveToNextCharacter()
    {
        // Strzałka w prawo - ogłoś znak na który przeszliśmy
        Task.Run(async () =>
        {
            await Task.Delay(50);
            _editableTextHandler.ReadCurrentCharacter();
        });
    }

    private void OnMoveToPreviousLine()
    {
        // Strzałka w górę - ogłoś linię na którą przeszliśmy
        Task.Run(async () =>
        {
            await Task.Delay(50);
            _editableTextHandler.ReadCurrentLine();
        });
    }

    private void OnMoveToNextLine()
    {
        // Strzałka w dół - ogłoś linię na którą przeszliśmy
        Task.Run(async () =>
        {
            await Task.Delay(50);
            _editableTextHandler.ReadCurrentLine();
        });
    }

    private void OnMoveToPreviousWord()
    {
        // Ctrl+strzałka w lewo - ogłoś słowo na które przeszliśmy
        Task.Run(async () =>
        {
            await Task.Delay(50);
            _editableTextHandler.ReadCurrentWord();
        });
    }

    private void OnMoveToNextWord()
    {
        // Ctrl+strzałka w prawo - ogłoś słowo na które przeszliśmy
        Task.Run(async () =>
        {
            await Task.Delay(50);
            _editableTextHandler.ReadCurrentWord();
        });
    }

    private void OnMoveToStart()
    {
        // Home - ogłoś pozycję
        Task.Run(async () =>
        {
            await Task.Delay(50);
            _speechManager.Speak("Początek");
        });
    }

    private void OnMoveToEnd()
    {
        // End - ogłoś pozycję
        Task.Run(async () =>
        {
            await Task.Delay(50);
            _speechManager.Speak("Koniec");
        });
    }

    private void OnReadCurrentChar()
    {
        _editableTextHandler.ReadCurrentCharacter();
    }

    private void OnReadCurrentWord()
    {
        _editableTextHandler.ReadCurrentWord();
    }

    private void OnReadCurrentLine()
    {
        _editableTextHandler.ReadCurrentLine();
    }

    private void OnReadPosition()
    {
        _editableTextHandler.ReadPosition();
    }

    // Terminal navigation handlers
    private string OnTerminalPreviousLine()
    {
        var terminal = _appModuleManager.GetTerminalModule();
        if (terminal == null)
            return "";

        string result = terminal.MoveToPreviousLine();
        _speechManager.Speak(result);
        return result;
    }

    private string OnTerminalNextLine()
    {
        var terminal = _appModuleManager.GetTerminalModule();
        if (terminal == null)
            return "";

        string result = terminal.MoveToNextLine();
        _speechManager.Speak(result);
        return result;
    }

    private string OnTerminalPreviousChar()
    {
        var terminal = _appModuleManager.GetTerminalModule();
        if (terminal == null)
            return "";

        string result = terminal.MoveToPreviousChar();
        _speechManager.Speak(result);
        return result;
    }

    private string OnTerminalNextChar()
    {
        var terminal = _appModuleManager.GetTerminalModule();
        if (terminal == null)
            return "";

        string result = terminal.MoveToNextChar();
        _speechManager.Speak(result);
        return result;
    }

    private string OnTerminalPreviousWord()
    {
        var terminal = _appModuleManager.GetTerminalModule();
        if (terminal == null)
            return "";

        string result = terminal.MoveToPreviousWord();
        _speechManager.Speak(result);
        return result;
    }

    private string OnTerminalNextWord()
    {
        var terminal = _appModuleManager.GetTerminalModule();
        if (terminal == null)
            return "";

        string result = terminal.MoveToNextWord();
        _speechManager.Speak(result);
        return result;
    }

    private string OnTerminalPreviousPage()
    {
        var terminal = _appModuleManager.GetTerminalModule();
        if (terminal == null)
            return "";

        string result = terminal.MoveToPreviousPage();
        _speechManager.Speak(result);
        return result;
    }

    private string OnTerminalNextPage()
    {
        var terminal = _appModuleManager.GetTerminalModule();
        if (terminal == null)
            return "";

        string result = terminal.MoveToNextPage();
        _speechManager.Speak(result);
        return result;
    }

    private string OnTerminalReadLine()
    {
        var terminal = _appModuleManager.GetTerminalModule();
        if (terminal == null)
            return "";

        string result = terminal.ReadCurrentLine();
        _speechManager.Speak(result);
        return result;
    }

    private void AnnounceElement(AutomationElement element, bool checkEdges)
    {
        var description = UIAutomationHelper.GetElementDescription(element);
        Console.WriteLine($"Element: {description}");

        // Speak the description with TTS
        _speechManager.Speak(description);

        // Play appropriate sound based on element type
        if (UIAutomationHelper.IsListItem(element))
        {
            float position = UIAutomationHelper.GetListItemPosition(element);
            _soundManager.PlayListItem(position);
            Console.WriteLine($"  Pozycja na liście: {position:P0}");
            
            // Play edge sound if at beginning or end of list
            bool isAtStart = UIAutomationHelper.IsAtEdge(element, false); // no previous
            bool isAtEnd = UIAutomationHelper.IsAtEdge(element, true); // no next
            
            if (isAtStart || isAtEnd)
            {
                _soundManager.PlayEdge();
                if (isAtStart && isAtEnd)
                {
                    Console.WriteLine("  [Jedyny element na liście]");
                }
                else if (isAtStart)
                {
                    Console.WriteLine("  [Początek listy]");
                }
                else
                {
                    Console.WriteLine("  [Koniec listy]");
                }
            }
        }
        else if (UIAutomationHelper.IsButton(element))
        {
            // Buttons don't auto-play, only on click
            Console.WriteLine("  [Przycisk]");
        }
        else
        {
            // Default cursor sound for other elements
            if (!checkEdges) // Don't play cursor sound if we just navigated
                _soundManager.PlayCursor();
        }
    }

    private AutomationElement? GetContainingWindow(AutomationElement? element)
    {
        var current = element;
        while (current != null)
        {
            if (UIAutomationHelper.IsWindow(current))
                return current;
            current = UIAutomationHelper.GetParent(current);
        }
        return null;
    }

    private string GetWindowTitle(AutomationElement? window)
    {
        if (window == null)
            return "Nieznane okno";
        
        try
        {
            return window.Current.Name;
        }
        catch
        {
            return "Nieznane okno";
        }
    }

    /// <summary>
    /// Przełącza echo klawiatury i zwraca komunikat
    /// </summary>
    public string CycleKeyboardEcho()
    {
        _keyboardEchoMode = _keyboardEchoMode.Next();
        return $"Echo klawiszy: {_keyboardEchoMode.GetPolishName()}";
    }

    private void RegisterCustomGestures()
    {
        // Insert+Ctrl+Space - Toggle virtual cursor (TCE)
        _gestureManager.RegisterGesture("insert+ctrl+space", System.Windows.Forms.Keys.Space,
            () =>
            {
                string message = _browseModeHandler.ToggleVirtualCursor();
                _speechManager.Speak(message);
            },
            "Przełącz kursor TCE",
            "Włącza lub wyłącza kursor trybu czytania ekranu",
            "Nawigacja");

        // Insert+2 - Cycle keyboard echo
        _gestureManager.RegisterGesture("insert+2", System.Windows.Forms.Keys.D2,
            () =>
            {
                string message = CycleKeyboardEcho();
                _speechManager.Speak(message);
            },
            "Przełącz echo klawiszy",
            "Przełącza między trybami echa klawiatury: znaki, słowa, słowa i znaki, brak",
            "Mowa");

        // Insert+N - Show screen reader menu
        _gestureManager.RegisterGesture("insert+n", System.Windows.Forms.Keys.N,
            () => OnShowMenu(),
            "Menu czytnika ekranu",
            "Otwiera menu czytnika ekranu",
            "System");

        // Insert+T - Read window title
        _gestureManager.RegisterGesture("insert+t", System.Windows.Forms.Keys.T,
            () =>
            {
                var window = GetContainingWindow(_currentElement);
                if (window != null)
                {
                    var title = GetWindowTitle(window);
                    _speechManager.Speak($"Okno: {title}");
                }
                else
                {
                    _speechManager.Speak("Nie znaleziono okna");
                }
            },
            "Odczytaj tytuł okna",
            "Ogłasza tytuł aktywnego okna",
            "Nawigacja");

        // Insert+Ctrl+T - Read current element type
        _gestureManager.RegisterGesture("insert+ctrl+t", System.Windows.Forms.Keys.T,
            () =>
            {
                if (_currentElement != null)
                {
                    var controlType = UIAutomationHelper.GetPolishControlType(_currentElement.Current.ControlType);
                    _speechManager.Speak($"Typ: {controlType}");
                }
            },
            "Odczytaj typ elementu",
            "Ogłasza typ bieżącego elementu",
            "Informacje");

        Console.WriteLine("ScreenReaderEngine: Zarejestrowano niestandardowe gesty");
    }

    public void Stop()
    {
        Console.WriteLine("Zatrzymywanie Czytnika Ekranu...");
        
        _focusTracker.Stop();
        _keyboardHook.Stop();
        _soundManager.Stop();
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        Stop();

        // Dispose nowych komponentów
        _appModuleManager.Dispose();
        _browseModeHandler.Dispose();
        _editableTextHandler.Dispose();

        _soundManager.Dispose();
        _focusTracker.Dispose();
        _keyboardHook.Dispose();
        _speechManager.Dispose();
        _dialogMonitor?.Dispose();

        Instance = null;
        _disposed = true;
    }
}
