using System.Windows.Automation;
using System.Windows.Automation.Text;

namespace ScreenReader;

public class EditFieldNavigator
{
    private AutomationElement? _currentEdit;
    private TextPattern? _textPattern;
    private readonly SpeechManager _speechManager;
    
    // Polski alfabet fonetyczny
    private static readonly Dictionary<char, string> PhoneticAlphabet = new()
    {
        {'a', "Anna"}, {'ą', "Ąbecadło"}, {'b', "Barbara"}, {'c', "Celina"},
        {'ć', "Ćma"}, {'d', "Dorota"}, {'e', "Ewa"}, {'ę', "Ęby"},
        {'f', "Franciszek"}, {'g', "Genowefa"}, {'h', "Henryk"}, {'i', "Irena"},
        {'j', "Janina"}, {'k', "Katarzyna"}, {'l', "Leon"}, {'ł', "Łódź"},
        {'m', "Maria"}, {'n', "Natalia"}, {'ń', "Ńwieboda"}, {'o', "Olga"},
        {'ó', "Ósemka"}, {'p', "Paweł"}, {'q', "Quebec"}, {'r', "Roman"},
        {'s', "Stefan"}, {'ś', "Świerk"}, {'t', "Tadeusz"}, {'u', "Urszula"},
        {'v', "Violetta"}, {'w', "Wanda"}, {'x', "Xawery"}, {'y', "Ypsylon"},
        {'z', "Zofia"}, {'ź', "Źrebak"}, {'ż', "Żaba"},
        {' ', "spacja"}, {'.', "kropka"}, {',', "przecinek"},
        {'!', "wykrzyknik"}, {'?', "pytajnik"}, {'-', "myślnik"},
        {'_', "podkreślenie"}, {'/', "ukośnik"}, {'\\', "ukośnik wsteczny"},
        {'@', "małpa"}, {'#', "hash"}, {'$', "dolar"}, {'%', "procent"},
        {'^', "daszek"}, {'&', "ampersand"}, {'*', "gwiazdka"},
        {'(', "lewy nawias"}, {')', "prawy nawias"}, {'[', "lewy kwadratowy"},
        {']', "prawy kwadratowy"}, {'{', "lewy klamrowy"}, {'}', "prawy klamrowy"},
        {'<', "mniejsze"}, {'>', "większe"}, {'=', "równe"}, {'+', "plus"},
        {':', "dwukropek"}, {';', "średnik"}, {'\"', "cudzysłów"}, {'\'', "apostrof"}
    };

    public EditFieldNavigator(SpeechManager speechManager)
    {
        _speechManager = speechManager;
    }

    public bool IsInEditField(AutomationElement? element)
    {
        if (element == null)
            return false;

        try
        {
            var controlType = element.Current.ControlType;
            return controlType == ControlType.Edit || 
                   controlType == ControlType.Document;
        }
        catch
        {
            return false;
        }
    }

    public void SetCurrentEdit(AutomationElement? element)
    {
        if (element == null || !IsInEditField(element))
        {
            _currentEdit = null;
            _textPattern = null;
            return;
        }

        try
        {
            _currentEdit = element;
            
            if (element.TryGetCurrentPattern(TextPattern.Pattern, out object? pattern))
            {
                _textPattern = pattern as TextPattern;
                Console.WriteLine("EditFieldNavigator: TextPattern wykryty");
            }
            else
            {
                _textPattern = null;
                Console.WriteLine("EditFieldNavigator: Brak TextPattern");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"EditFieldNavigator: Błąd ustawienia pola: {ex.Message}");
            _textPattern = null;
        }
    }

    public string GetCurrentCharacter()
    {
        if (_textPattern == null)
            return "";

        try
        {
            var selection = _textPattern.GetSelection();
            if (selection.Length == 0)
                return "";

            var caretRange = selection[0];
            var charRange = caretRange.Clone();
            
            // Rozszerz o jeden znak w prawo
            int moved = charRange.MoveEndpointByUnit(
                TextPatternRangeEndpoint.End,
                TextUnit.Character,
                1);

            if (moved == 0)
                return "koniec tekstu";

            string text = charRange.GetText(-1);
            if (string.IsNullOrEmpty(text))
                return "pusty";

            char ch = text[0];
            return GetCharacterDescription(ch);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"EditFieldNavigator: Błąd odczytu znaku: {ex.Message}");
            return "";
        }
    }

    public string GetCurrentWord()
    {
        if (_textPattern == null)
            return "";

        try
        {
            var selection = _textPattern.GetSelection();
            if (selection.Length == 0)
                return "";

            var caretRange = selection[0].Clone();
            caretRange.ExpandToEnclosingUnit(TextUnit.Word);
            
            string word = caretRange.GetText(-1);
            return string.IsNullOrWhiteSpace(word) ? "puste" : word.Trim();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"EditFieldNavigator: Błąd odczytu słowa: {ex.Message}");
            return "";
        }
    }

    public string GetCurrentLine()
    {
        if (_textPattern == null)
            return "";

        try
        {
            var selection = _textPattern.GetSelection();
            if (selection.Length == 0)
                return "";

            var caretRange = selection[0].Clone();
            caretRange.ExpandToEnclosingUnit(TextUnit.Line);
            
            string line = caretRange.GetText(-1);
            return string.IsNullOrWhiteSpace(line) ? "pusta linia" : line.Trim();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"EditFieldNavigator: Błąd odczytu linii: {ex.Message}");
            return "";
        }
    }

    public void MoveByCharacter(int count)
    {
        if (_textPattern == null)
        {
            _speechManager.Speak("Nie można nawigować");
            return;
        }

        try
        {
            var selection = _textPattern.GetSelection();
            if (selection.Length == 0)
                return;

            var range = selection[0].Clone();
            int moved = range.Move(TextUnit.Character, count);
            
            if (moved == 0)
            {
                _speechManager.Speak(count > 0 ? "Koniec" : "Początek");
                return;
            }

            range.Select();
            
            // Odczytaj nowy znak
            string ch = GetCurrentCharacter();
            if (!string.IsNullOrEmpty(ch))
                _speechManager.Speak(ch);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"EditFieldNavigator: Błąd nawigacji znakowej: {ex.Message}");
            _speechManager.Speak("Błąd nawigacji");
        }
    }

    public void MoveByWord(int count)
    {
        if (_textPattern == null)
        {
            _speechManager.Speak("Nie można nawigować");
            return;
        }

        try
        {
            var selection = _textPattern.GetSelection();
            if (selection.Length == 0)
                return;

            var range = selection[0].Clone();
            int moved = range.Move(TextUnit.Word, count);
            
            if (moved == 0)
            {
                _speechManager.Speak(count > 0 ? "Koniec" : "Początek");
                return;
            }

            range.Select();
            
            // Odczytaj nowe słowo
            string word = GetCurrentWord();
            if (!string.IsNullOrEmpty(word))
                _speechManager.Speak(word);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"EditFieldNavigator: Błąd nawigacji słownej: {ex.Message}");
            _speechManager.Speak("Błąd nawigacji");
        }
    }

    public void MoveToStart()
    {
        if (_textPattern == null)
        {
            _speechManager.Speak("Nie można nawigować");
            return;
        }

        try
        {
            var docRange = _textPattern.DocumentRange;
            var startRange = docRange.Clone();
            startRange.MoveEndpointByRange(TextPatternRangeEndpoint.End, startRange, TextPatternRangeEndpoint.Start);
            startRange.Select();
            
            _speechManager.Speak("Początek");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"EditFieldNavigator: Błąd przejścia do początku: {ex.Message}");
        }
    }

    public void MoveToEnd()
    {
        if (_textPattern == null)
        {
            _speechManager.Speak("Nie można nawigować");
            return;
        }

        try
        {
            var docRange = _textPattern.DocumentRange;
            var endRange = docRange.Clone();
            endRange.MoveEndpointByRange(TextPatternRangeEndpoint.Start, endRange, TextPatternRangeEndpoint.End);
            endRange.Select();
            
            _speechManager.Speak("Koniec");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"EditFieldNavigator: Błąd przejścia do końca: {ex.Message}");
        }
    }

    public string GetPositionInfo()
    {
        if (_textPattern == null)
            return "Brak informacji o pozycji";

        try
        {
            var docRange = _textPattern.DocumentRange;
            var selection = _textPattern.GetSelection();
            
            if (selection.Length == 0)
                return "Brak zaznaczenia";

            var caretRange = selection[0];
            
            // Oblicz pozycję znaku od początku dokumentu
            int charPosition = caretRange.CompareEndpoints(
                TextPatternRangeEndpoint.Start,
                docRange,
                TextPatternRangeEndpoint.Start);

            // Zlicz linie do kursora
            var lineRange = docRange.Clone();
            lineRange.MoveEndpointByRange(TextPatternRangeEndpoint.End, caretRange, TextPatternRangeEndpoint.Start);
            
            string textToCursor = lineRange.GetText(-1);
            int lineNumber = textToCursor.Split('\n').Length;

            return $"Linia {lineNumber}, znak {charPosition + 1}";
        }
        catch (Exception ex)
        {
            Console.WriteLine($"EditFieldNavigator: Błąd pozycji: {ex.Message}");
            return "Błąd odczytu pozycji";
        }
    }

    private string GetCharacterDescription(char c)
    {
        if (PhoneticAlphabet.TryGetValue(char.ToLower(c), out var phonetic))
        {
            if (char.IsUpper(c))
                return $"Wielka {c}, {phonetic}";
            return $"{c}, {phonetic}";
        }

        if (char.IsDigit(c))
            return $"cyfra {c}";

        if (char.IsWhiteSpace(c))
            return "spacja";

        return c.ToString();
    }
}
