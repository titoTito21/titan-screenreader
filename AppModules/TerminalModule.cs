using System.Text;
using System.Windows.Automation;

namespace ScreenReader.AppModules;

/// <summary>
/// Tryb nawigacji NumPad dla terminali
/// </summary>
public enum TerminalNavMode
{
    /// <summary>Nawigacja obiektowa (domyślna)</summary>
    Object,

    /// <summary>Nawigacja tekstowa (dla terminali)</summary>
    Text
}

/// <summary>
/// Moduł dla terminali konsolowych (cmd, PowerShell, Windows Terminal)
///
/// Obsługuje specjalną nawigację NumPad:
/// - 2: następna linia tekstu
/// - 8: poprzednia linia tekstu
/// - 4: poprzedni znak
/// - 6: następny znak
/// - 1: poprzedni wyraz
/// - 3: następny wyraz
/// - 7: poprzednia strona (Page Up)
/// - 9: następna strona (Page Down)
/// - 5: odczytaj bieżącą linię
/// </summary>
public class TerminalModule : AppModuleBase
{
    private string[] _screenBuffer = Array.Empty<string>();
    private int _currentLine;
    private int _currentColumn;
    private AutomationElement? _terminalElement;
    private DateTime _lastBufferUpdate = DateTime.MinValue;
    private const int BufferCacheMs = 100;

    /// <summary>Tryb nawigacji</summary>
    public TerminalNavMode NavMode { get; set; } = TerminalNavMode.Text;

    /// <summary>Zdarzenie odczytu tekstu</summary>
    public event Action<string>? TextRead;

    public TerminalModule(string processName) : base(processName)
    {
    }

    public override void OnGainFocus(AutomationElement element)
    {
        base.OnGainFocus(element);
        _terminalElement = FindTerminalContent(element);
        RefreshBuffer();
        Console.WriteLine($"TerminalModule: Aktywny ({ProcessName})");
    }

    public override void OnFocusChanged(AutomationElement element)
    {
        base.OnFocusChanged(element);

        // Odśwież bufor przy zmianie fokusu wewnątrz terminala
        var newTerminal = FindTerminalContent(element);
        if (newTerminal != null)
        {
            _terminalElement = newTerminal;
        }
    }

    public override void OnLoseFocus()
    {
        base.OnLoseFocus();
        _terminalElement = null;
    }

    /// <summary>
    /// Znajduje element z treścią terminala
    /// </summary>
    private AutomationElement? FindTerminalContent(AutomationElement element)
    {
        try
        {
            // Dla Windows Terminal - szukaj elementu Document lub Text
            var walker = TreeWalker.ControlViewWalker;
            var current = element;

            // Szukaj w górę drzewa
            while (current != null)
            {
                var controlType = current.Current.ControlType;

                // Terminal ma zwykle Document lub Edit jako główny element treści
                if (controlType == ControlType.Document ||
                    controlType == ControlType.Edit)
                {
                    return current;
                }

                // Dla cmd.exe/powershell.exe - szukaj okna z tekstem
                if (controlType == ControlType.Window)
                {
                    // Spróbuj znaleźć child z TextPattern
                    var child = walker.GetFirstChild(current);
                    while (child != null)
                    {
                        if (child.TryGetCurrentPattern(TextPattern.Pattern, out _))
                        {
                            return child;
                        }
                        child = walker.GetNextSibling(child);
                    }
                }

                current = walker.GetParent(current);
            }

            // Fallback - użyj samego elementu
            return element;
        }
        catch
        {
            return element;
        }
    }

    /// <summary>
    /// Odświeża bufor ekranu terminala
    /// </summary>
    public void RefreshBuffer()
    {
        var now = DateTime.Now;
        if ((now - _lastBufferUpdate).TotalMilliseconds < BufferCacheMs)
            return;

        _lastBufferUpdate = now;

        try
        {
            if (_terminalElement == null)
                return;

            string text = GetTerminalText();
            if (string.IsNullOrEmpty(text))
                return;

            // Podziel na linie
            _screenBuffer = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            // Ogranicz pozycję do zakresu
            _currentLine = Math.Clamp(_currentLine, 0, Math.Max(0, _screenBuffer.Length - 1));

            if (_screenBuffer.Length > 0 && _currentLine < _screenBuffer.Length)
            {
                _currentColumn = Math.Clamp(_currentColumn, 0, Math.Max(0, _screenBuffer[_currentLine].Length - 1));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"TerminalModule: Błąd odświeżania bufora: {ex.Message}");
        }
    }

    /// <summary>
    /// Pobiera tekst z terminala
    /// </summary>
    private string GetTerminalText()
    {
        if (_terminalElement == null)
            return "";

        try
        {
            // Próbuj TextPattern
            if (_terminalElement.TryGetCurrentPattern(TextPattern.Pattern, out var textPattern))
            {
                return ((TextPattern)textPattern).DocumentRange.GetText(-1);
            }

            // Próbuj ValuePattern
            if (_terminalElement.TryGetCurrentPattern(ValuePattern.Pattern, out var valuePattern))
            {
                return ((ValuePattern)valuePattern).Current.Value;
            }

            // Fallback - nazwa elementu
            return _terminalElement.Current.Name ?? "";
        }
        catch
        {
            return "";
        }
    }

    /// <summary>
    /// Przechodzi do następnej linii (NumPad 2)
    /// </summary>
    public string MoveToNextLine()
    {
        RefreshBuffer();

        if (_screenBuffer.Length == 0)
            return "Pusty bufor";

        if (_currentLine < _screenBuffer.Length - 1)
        {
            _currentLine++;
            _currentColumn = 0;
            return GetCurrentLineText();
        }

        return "Koniec bufora";
    }

    /// <summary>
    /// Przechodzi do poprzedniej linii (NumPad 8)
    /// </summary>
    public string MoveToPreviousLine()
    {
        RefreshBuffer();

        if (_screenBuffer.Length == 0)
            return "Pusty bufor";

        if (_currentLine > 0)
        {
            _currentLine--;
            _currentColumn = 0;
            return GetCurrentLineText();
        }

        return "Początek bufora";
    }

    /// <summary>
    /// Przechodzi do następnego znaku (NumPad 6)
    /// </summary>
    public string MoveToNextChar()
    {
        RefreshBuffer();

        if (_screenBuffer.Length == 0 || _currentLine >= _screenBuffer.Length)
            return "Pusty bufor";

        var line = _screenBuffer[_currentLine];

        if (_currentColumn < line.Length - 1)
        {
            _currentColumn++;
            return GetCurrentCharText();
        }

        // Przejdź do następnej linii
        if (_currentLine < _screenBuffer.Length - 1)
        {
            _currentLine++;
            _currentColumn = 0;
            return GetCurrentCharText();
        }

        return "Koniec";
    }

    /// <summary>
    /// Przechodzi do poprzedniego znaku (NumPad 4)
    /// </summary>
    public string MoveToPreviousChar()
    {
        RefreshBuffer();

        if (_screenBuffer.Length == 0 || _currentLine >= _screenBuffer.Length)
            return "Pusty bufor";

        if (_currentColumn > 0)
        {
            _currentColumn--;
            return GetCurrentCharText();
        }

        // Przejdź do poprzedniej linii
        if (_currentLine > 0)
        {
            _currentLine--;
            var prevLine = _screenBuffer[_currentLine];
            _currentColumn = Math.Max(0, prevLine.Length - 1);
            return GetCurrentCharText();
        }

        return "Początek";
    }

    /// <summary>
    /// Przechodzi do następnego wyrazu (NumPad 3)
    /// </summary>
    public string MoveToNextWord()
    {
        RefreshBuffer();

        if (_screenBuffer.Length == 0 || _currentLine >= _screenBuffer.Length)
            return "Pusty bufor";

        var line = _screenBuffer[_currentLine];

        // Znajdź początek następnego słowa
        int pos = _currentColumn;

        // Pomiń bieżące słowo
        while (pos < line.Length && !char.IsWhiteSpace(line[pos]))
            pos++;

        // Pomiń białe znaki
        while (pos < line.Length && char.IsWhiteSpace(line[pos]))
            pos++;

        if (pos < line.Length)
        {
            _currentColumn = pos;
            return GetCurrentWordText();
        }

        // Przejdź do następnej linii
        if (_currentLine < _screenBuffer.Length - 1)
        {
            _currentLine++;
            _currentColumn = 0;

            // Pomiń białe znaki na początku linii
            var nextLine = _screenBuffer[_currentLine];
            while (_currentColumn < nextLine.Length && char.IsWhiteSpace(nextLine[_currentColumn]))
                _currentColumn++;

            return GetCurrentWordText();
        }

        return "Koniec";
    }

    /// <summary>
    /// Przechodzi do poprzedniego wyrazu (NumPad 1)
    /// </summary>
    public string MoveToPreviousWord()
    {
        RefreshBuffer();

        if (_screenBuffer.Length == 0 || _currentLine >= _screenBuffer.Length)
            return "Pusty bufor";

        var line = _screenBuffer[_currentLine];
        int pos = _currentColumn;

        // Pomiń białe znaki przed kursorem
        while (pos > 0 && char.IsWhiteSpace(line[pos - 1]))
            pos--;

        // Znajdź początek bieżącego słowa
        while (pos > 0 && !char.IsWhiteSpace(line[pos - 1]))
            pos--;

        if (pos > 0 || _currentColumn > 0)
        {
            _currentColumn = pos;
            return GetCurrentWordText();
        }

        // Przejdź do poprzedniej linii
        if (_currentLine > 0)
        {
            _currentLine--;
            var prevLine = _screenBuffer[_currentLine];
            _currentColumn = prevLine.Length;

            // Znajdź ostatnie słowo
            while (_currentColumn > 0 && char.IsWhiteSpace(prevLine[_currentColumn - 1]))
                _currentColumn--;
            while (_currentColumn > 0 && !char.IsWhiteSpace(prevLine[_currentColumn - 1]))
                _currentColumn--;

            return GetCurrentWordText();
        }

        return "Początek";
    }

    /// <summary>
    /// Przechodzi o stronę w górę (NumPad 7)
    /// </summary>
    public string MoveToPreviousPage()
    {
        RefreshBuffer();

        if (_screenBuffer.Length == 0)
            return "Pusty bufor";

        int pageSize = 25; // Typowa wysokość terminala
        _currentLine = Math.Max(0, _currentLine - pageSize);
        _currentColumn = 0;

        return $"Strona {(_currentLine / pageSize) + 1}, linia {_currentLine + 1}";
    }

    /// <summary>
    /// Przechodzi o stronę w dół (NumPad 9)
    /// </summary>
    public string MoveToNextPage()
    {
        RefreshBuffer();

        if (_screenBuffer.Length == 0)
            return "Pusty bufor";

        int pageSize = 25;
        _currentLine = Math.Min(_screenBuffer.Length - 1, _currentLine + pageSize);
        _currentColumn = 0;

        return $"Strona {(_currentLine / pageSize) + 1}, linia {_currentLine + 1}";
    }

    /// <summary>
    /// Odczytuje bieżącą linię (NumPad 5)
    /// </summary>
    public string ReadCurrentLine()
    {
        RefreshBuffer();
        return GetCurrentLineText();
    }

    /// <summary>
    /// Pobiera tekst bieżącej linii
    /// </summary>
    private string GetCurrentLineText()
    {
        if (_screenBuffer.Length == 0 || _currentLine >= _screenBuffer.Length)
            return "Pusta linia";

        var line = _screenBuffer[_currentLine];
        if (string.IsNullOrWhiteSpace(line))
            return "Pusta linia";

        return line;
    }

    /// <summary>
    /// Pobiera tekst bieżącego znaku z opisem
    /// </summary>
    private string GetCurrentCharText()
    {
        if (_screenBuffer.Length == 0 || _currentLine >= _screenBuffer.Length)
            return "Pusty";

        var line = _screenBuffer[_currentLine];
        if (_currentColumn >= line.Length)
            return "Koniec linii";

        char ch = line[_currentColumn];
        return GetCharDescription(ch);
    }

    /// <summary>
    /// Pobiera opis znaku
    /// </summary>
    private static string GetCharDescription(char ch)
    {
        return ch switch
        {
            ' ' => "spacja",
            '\t' => "tabulator",
            '.' => "kropka",
            ',' => "przecinek",
            ':' => "dwukropek",
            ';' => "średnik",
            '!' => "wykrzyknik",
            '?' => "pytajnik",
            '-' => "minus",
            '_' => "podkreślenie",
            '/' => "ukośnik",
            '\\' => "odwrotny ukośnik",
            '|' => "kreska pionowa",
            '@' => "małpa",
            '#' => "hash",
            '$' => "dolar",
            '%' => "procent",
            '&' => "ampersand",
            '*' => "gwiazdka",
            '(' => "nawias otwierający",
            ')' => "nawias zamykający",
            '[' => "nawias kwadratowy otwierający",
            ']' => "nawias kwadratowy zamykający",
            '{' => "nawias klamrowy otwierający",
            '}' => "nawias klamrowy zamykający",
            '<' => "mniejszy niż",
            '>' => "większy niż",
            '=' => "równa się",
            '+' => "plus",
            '\'' => "apostrof",
            '"' => "cudzysłów",
            '`' => "grawis",
            '~' => "tylda",
            '^' => "daszek",
            >= 'A' and <= 'Z' => $"duże {ch}",
            _ => ch.ToString()
        };
    }

    /// <summary>
    /// Pobiera tekst bieżącego wyrazu
    /// </summary>
    private string GetCurrentWordText()
    {
        if (_screenBuffer.Length == 0 || _currentLine >= _screenBuffer.Length)
            return "Pusty";

        var line = _screenBuffer[_currentLine];
        if (_currentColumn >= line.Length)
            return "Koniec linii";

        // Znajdź granice słowa
        int start = _currentColumn;
        int end = _currentColumn;

        // Początek słowa
        while (start > 0 && !char.IsWhiteSpace(line[start - 1]))
            start--;

        // Koniec słowa
        while (end < line.Length && !char.IsWhiteSpace(line[end]))
            end++;

        if (start == end)
            return "Spacja";

        return line.Substring(start, end - start);
    }

    /// <summary>
    /// Pobiera informacje o pozycji
    /// </summary>
    public string GetPositionInfo()
    {
        RefreshBuffer();
        return $"Linia {_currentLine + 1} z {_screenBuffer.Length}, kolumna {_currentColumn + 1}";
    }

    public override bool ShouldUseVirtualBuffer(AutomationElement element)
    {
        // Terminale nie używają wirtualnego bufora
        return false;
    }
}

/// <summary>
/// Moduł dla Windows Terminal
/// </summary>
public class WindowsTerminalModule : TerminalModule
{
    public WindowsTerminalModule() : base("WindowsTerminal")
    {
    }
}

/// <summary>
/// Moduł dla Command Prompt (cmd.exe)
/// </summary>
public class CmdModule : TerminalModule
{
    public CmdModule() : base("cmd")
    {
    }
}

/// <summary>
/// Moduł dla PowerShell
/// </summary>
public class PowerShellModule : TerminalModule
{
    public PowerShellModule() : base("powershell")
    {
    }
}

/// <summary>
/// Moduł dla PowerShell 7+ (pwsh.exe)
/// </summary>
public class PwshModule : TerminalModule
{
    public PwshModule() : base("pwsh")
    {
    }
}
