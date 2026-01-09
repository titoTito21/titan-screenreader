using System.Windows.Automation;
using ScreenReader.Settings;

namespace ScreenReader.Hints;

/// <summary>
/// Zarządza podpowiedziami dla kontrolek.
/// Podpowiedzi są odczytywane 2 sekundy po zatrzymaniu się na kontrolce.
/// </summary>
public class HintManager : IDisposable
{
    private readonly SpeechManager _speechManager;
    private readonly SettingsManager _settings;
    private System.Threading.Timer? _hintTimer;
    private AutomationElement? _currentElement;
    private ControlType? _currentControlType;
    private bool _disposed;
    private const int HINT_DELAY_MS = 2000; // 2 sekundy

    // Słownik podpowiedzi dla różnych typów kontrolek
    private static readonly Dictionary<string, string> ControlHints = new()
    {
        // Podstawowe kontrolki
        { "Button", "Naciśnij Enter lub spację, aby aktywować" },
        { "CheckBox", "Aby zaznaczyć lub odznaczyć, naciśnij spację" },
        { "RadioButton", "Naciśnij spację, aby wybrać tę opcję" },
        { "Edit", "Zacznij pisać, aby edytować" },
        { "Document", "Zacznij pisać, aby edytować" },
        { "Text", "To jest tekst statyczny" },

        // Listy i drzewa
        { "List", "Użyj strzałek góra/dół, aby nawigować po liście" },
        { "ListItem", "Naciśnij Enter, aby aktywować element" },
        { "Tree", "Użyj strzałek do nawigacji. Prawo rozwija, lewo zwija" },
        { "TreeItem", "Naciśnij prawo, aby rozwinąć. Lewo, aby zwinąć" },

        // Combo i menu
        { "ComboBox", "Naciśnij Alt plus strzałka w dół, aby rozwinąć listę" },
        { "Menu", "Użyj strzałek do nawigacji po menu" },
        { "MenuItem", "Naciśnij Enter, aby aktywować element menu" },
        { "MenuBar", "Użyj strzałek lewo/prawo do nawigacji po menu" },

        // Zakładki
        { "Tab", "Użyj Control plus Tab, aby przełączać zakładki" },
        { "TabItem", "Naciśnij Enter lub spację, aby wybrać zakładkę" },

        // Paski narzędzi i statusu
        { "ToolBar", "Użyj Tab i strzałek do nawigacji po pasku narzędzi" },
        { "StatusBar", "Aby przełączyć do listy aplikacji, naciśnij Tab" },

        // Suwaki i spinboxy
        { "Slider", "Użyj strzałek lewo/prawo lub góra/dół, aby zmienić wartość" },
        { "Spinner", "Użyj strzałek góra/dół, aby zmienić wartość" },

        // Linki i obrazy
        { "Hyperlink", "Naciśnij Enter, aby otworzyć link" },
        { "Image", "To jest obraz" },

        // Tabele
        { "Table", "Użyj Control plus strzałek do nawigacji po tabeli" },
        { "DataGrid", "Użyj strzałek do nawigacji po siatce danych" },

        // Okna i panele
        { "Window", "To jest okno aplikacji" },
        { "Pane", "To jest panel" },
        { "Group", "To jest grupa kontrolek" },

        // Paski przewijania
        { "ScrollBar", "Użyj strzałek lub Page Up/Down do przewijania" },

        // Nagłówki
        { "Header", "To jest nagłówek" },
        { "HeaderItem", "Kliknij, aby posortować kolumnę" },

        // Postęp
        { "ProgressBar", "To jest pasek postępu" },

        // Separatory
        { "Separator", "To jest separator" },

        // Miniaturki
        { "Thumb", "Przeciągnij, aby zmienić wartość" },

        // Kalendarze
        { "Calendar", "Użyj strzałek do nawigacji po kalendarzu" },

        // Niestandardowe
        { "Custom", "To jest kontrolka niestandardowa" },
    };

    // Podpowiedzi specyficzne dla TCE/Titan
    private static readonly Dictionary<string, string> TCEHints = new()
    {
        { "AppList", "Aby przełączyć na pasek stanu, naciśnij Tab. Aby przełączyć między listami, naciśnij Control plus Tab" },
        { "GameList", "Aby przełączyć na pasek stanu, naciśnij Tab. Aby przełączyć między listami, naciśnij Control plus Tab" },
        { "StatusBar", "Aby przełączyć do listy aplikacji, naciśnij Tab" },
    };

    public HintManager(SpeechManager speechManager)
    {
        _speechManager = speechManager;
        _settings = SettingsManager.Instance;
    }

    /// <summary>
    /// Ustawia bieżący element i resetuje timer podpowiedzi
    /// </summary>
    public void SetCurrentElement(AutomationElement? element, bool isTCEProcess = false)
    {
        // Zatrzymaj poprzedni timer
        _hintTimer?.Dispose();
        _hintTimer = null;

        _currentElement = element;

        if (element == null || !_settings.SpeakHints)
            return;

        try
        {
            _currentControlType = element.Current.ControlType;

            // Uruchom timer podpowiedzi
            _hintTimer = new System.Threading.Timer(
                OnHintTimerElapsed,
                isTCEProcess,
                HINT_DELAY_MS,
                Timeout.Infinite);
        }
        catch (ElementNotAvailableException)
        {
            // Element już niedostępny
        }
    }

    /// <summary>
    /// Anuluje oczekującą podpowiedź
    /// </summary>
    public void CancelHint()
    {
        _hintTimer?.Dispose();
        _hintTimer = null;
    }

    /// <summary>
    /// Callback timera - odczytuje podpowiedź
    /// </summary>
    private void OnHintTimerElapsed(object? state)
    {
        if (_currentElement == null || !_settings.SpeakHints)
            return;

        bool isTCE = state is bool b && b;

        try
        {
            string? hint = GetHintForElement(_currentElement, isTCE);

            if (!string.IsNullOrEmpty(hint))
            {
                // Odczytaj podpowiedź (nie przerywaj bieżącej mowy)
                _speechManager.Speak(hint, interrupt: false);
            }
        }
        catch (ElementNotAvailableException)
        {
            // Element już niedostępny
        }
        catch (Exception ex)
        {
            Console.WriteLine($"HintManager: Błąd odczytu podpowiedzi - {ex.Message}");
        }
    }

    /// <summary>
    /// Pobiera podpowiedź dla danego elementu
    /// </summary>
    private string? GetHintForElement(AutomationElement element, bool isTCE)
    {
        try
        {
            var controlType = element.Current.ControlType;
            string typeName = controlType.ProgrammaticName.Replace("ControlType.", "");

            // Sprawdź specjalne podpowiedzi TCE
            if (isTCE)
            {
                string? name = element.Current.Name?.ToLowerInvariant() ?? "";
                string? className = element.Current.ClassName?.ToLowerInvariant() ?? "";

                // Lista aplikacji lub gier w TCE
                if ((name.Contains("aplikacj") || name.Contains("gier") || name.Contains("gry")) &&
                    (typeName == "List" || typeName == "ListItem"))
                {
                    return TCEHints.GetValueOrDefault("AppList");
                }

                // Pasek stanu w TCE
                if (typeName == "StatusBar" || name.Contains("pasek stanu") || className.Contains("statusbar"))
                {
                    return TCEHints.GetValueOrDefault("StatusBar");
                }
            }

            // Standardowe podpowiedzi
            return ControlHints.GetValueOrDefault(typeName);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Pobiera podpowiedź dla danego typu kontrolki (bez elementu)
    /// </summary>
    public static string? GetHintForControlType(string controlTypeName)
    {
        return ControlHints.GetValueOrDefault(controlTypeName);
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _hintTimer?.Dispose();
        _hintTimer = null;
        _disposed = true;
    }
}
