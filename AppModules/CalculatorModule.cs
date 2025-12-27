using System.Windows.Automation;

namespace ScreenReader.AppModules;

/// <summary>
/// Moduł dla aplikacji Kalkulator Windows
/// Zapewnia ulepszone odczytywanie wyników obliczeń
/// </summary>
public class CalculatorModule : UWPModule
{
    private string? _lastResult;
    private string? _lastExpression;

    public CalculatorModule() : base("Calculator")
    {
    }

    public override void OnGainFocus(AutomationElement element)
    {
        base.OnGainFocus(element);
        Console.WriteLine("CalculatorModule: Kalkulator aktywny");

        // Odczytaj aktualny wynik przy wejściu
        TryReadCurrentResult(element);
    }

    public override void OnFocusChanged(AutomationElement element)
    {
        base.OnFocusChanged(element);

        // Sprawdź czy fokus jest na wyświetlaczu
        if (IsCalculatorDisplay(element))
        {
            AnnounceResult(element);
        }
        else if (IsCalculatorButton(element))
        {
            // Dla przycisków odczytaj nazwę
            AnnounceButton(element);
        }
    }

    /// <summary>
    /// Sprawdza czy element jest wyświetlaczem kalkulatora
    /// </summary>
    private bool IsCalculatorDisplay(AutomationElement element)
    {
        try
        {
            string automationId = element.Current.AutomationId;
            return automationId.Contains("Result", StringComparison.OrdinalIgnoreCase) ||
                   automationId.Contains("Display", StringComparison.OrdinalIgnoreCase) ||
                   automationId.Contains("CalculatorResults", StringComparison.OrdinalIgnoreCase) ||
                   automationId.Contains("Expression", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Sprawdza czy element jest przyciskiem kalkulatora
    /// </summary>
    private bool IsCalculatorButton(AutomationElement element)
    {
        try
        {
            return element.Current.ControlType == ControlType.Button;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Ogłasza wynik kalkulatora
    /// </summary>
    private void AnnounceResult(AutomationElement element)
    {
        try
        {
            string name = element.Current.Name;
            string automationId = element.Current.AutomationId;

            // Parsuj wynik - usuń "Wyświetlacz to" itp.
            string result = ParseCalculatorResult(name);

            // Sprawdź czy wynik się zmienił
            if (!string.IsNullOrEmpty(result) && result != _lastResult)
            {
                _lastResult = result;
                Console.WriteLine($"CalculatorModule: Wynik = {result}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"CalculatorModule: Błąd odczytu wyniku: {ex.Message}");
        }
    }

    /// <summary>
    /// Ogłasza przycisk kalkulatora
    /// </summary>
    private void AnnounceButton(AutomationElement element)
    {
        try
        {
            string name = element.Current.Name;
            string automationId = element.Current.AutomationId;

            // Mapuj nazwy przycisków na polskie
            string buttonText = MapButtonName(name, automationId);

            Console.WriteLine($"CalculatorModule: Przycisk {buttonText}");
        }
        catch
        {
            // Ignore
        }
    }

    /// <summary>
    /// Parsuje wynik z wyświetlacza kalkulatora
    /// </summary>
    private string ParseCalculatorResult(string displayText)
    {
        if (string.IsNullOrEmpty(displayText))
            return "";

        // Usuń prefiks "Wyświetlacz to" lub "Display is"
        var prefixes = new[] { "Wyświetlacz to ", "Display is ", "Wynik ", "Result " };
        foreach (var prefix in prefixes)
        {
            if (displayText.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                displayText = displayText.Substring(prefix.Length);
                break;
            }
        }

        return displayText.Trim();
    }

    /// <summary>
    /// Mapuje nazwy przycisków na polskie
    /// </summary>
    private string MapButtonName(string name, string automationId)
    {
        // Mapowanie operatorów
        return automationId switch
        {
            "plusButton" => "plus",
            "minusButton" => "minus",
            "multiplyButton" => "razy",
            "divideButton" => "dzielone przez",
            "equalButton" => "równa się",
            "clearButton" => "wyczyść",
            "clearEntryButton" => "wyczyść wpis",
            "backSpaceButton" => "cofnij",
            "negateButton" => "zmień znak",
            "decimalSeparatorButton" => "przecinek",
            "percentButton" => "procent",
            "squareRootButton" => "pierwiastek kwadratowy",
            "invertButton" => "jeden przez x",
            "squareButton" => "do kwadratu",
            "memoryAdd" => "dodaj do pamięci",
            "memorySubtract" => "odejmij z pamięci",
            "memoryRecall" => "przywołaj pamięć",
            "memoryClear" => "wyczyść pamięć",
            "memoryStore" => "zapisz w pamięci",
            _ => name
        };
    }

    /// <summary>
    /// Próbuje odczytać aktualny wynik przy wejściu do aplikacji
    /// </summary>
    private void TryReadCurrentResult(AutomationElement root)
    {
        try
        {
            // Szukaj wyświetlacza wyniku
            var resultElement = FindResultDisplay(root);
            if (resultElement != null)
            {
                AnnounceResult(resultElement);
            }
        }
        catch
        {
            // Ignore
        }
    }

    /// <summary>
    /// Znajduje element wyświetlacza wyniku
    /// </summary>
    private AutomationElement? FindResultDisplay(AutomationElement root)
    {
        try
        {
            // Szukaj po AutomationId
            var condition = new PropertyCondition(
                AutomationElement.AutomationIdProperty,
                "CalculatorResults");

            var element = root.FindFirst(TreeScope.Descendants, condition);
            if (element != null)
                return element;

            // Alternatywna metoda - szukaj Text z "Result"
            var textCondition = new PropertyCondition(
                AutomationElement.ControlTypeProperty,
                ControlType.Text);

            var elements = root.FindAll(TreeScope.Descendants, textCondition);
            foreach (AutomationElement el in elements)
            {
                try
                {
                    if (el.Current.AutomationId.Contains("Result"))
                        return el;
                }
                catch
                {
                    continue;
                }
            }
        }
        catch
        {
            // Ignore
        }

        return null;
    }

    public override void CustomizeElement(AutomationElement element, ref string name, ref string role)
    {
        base.CustomizeElement(element, ref name, ref role);

        try
        {
            // Dla wyświetlacza pokazuj sam wynik
            if (IsCalculatorDisplay(element))
            {
                name = ParseCalculatorResult(name);
                role = "wynik";
            }
            else if (IsCalculatorButton(element))
            {
                name = MapButtonName(name, element.Current.AutomationId);
                role = "przycisk";
            }
        }
        catch
        {
            // Ignore
        }
    }
}
