using System.Windows.Automation;

namespace ScreenReader;

public class UIAutomationHelper
{
    public static AutomationElement? GetFocusedElement()
    {
        try
        {
            return AutomationElement.FocusedElement;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting focused element: {ex.Message}");
            return null;
        }
    }

    public static string GetElementDescription(AutomationElement? element)
    {
        if (element == null)
            return "Brak elementu";

        try
        {
            var name = element.Current.Name;
            var controlType = element.Current.ControlType.ProgrammaticName.Replace("ControlType.", "");
            var controlTypePolish = TranslateControlType(controlType);
            var value = GetElementValue(element);
            var helpText = element.Current.HelpText;

            var description = string.IsNullOrWhiteSpace(name) ? controlTypePolish : $"{name}, {controlTypePolish}";

            if (!string.IsNullOrWhiteSpace(value))
            {
                description += $", {value}";
            }

            if (!string.IsNullOrWhiteSpace(helpText))
            {
                description += $", {helpText}";
            }

            return description;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd odczytu elementu: {ex.Message}");
            return "Nieznany element";
        }
    }

    /// <summary>
    /// Pobiera polską nazwę typu kontrolki
    /// </summary>
    public static string GetPolishControlType(ControlType controlType)
    {
        var typeName = controlType.ProgrammaticName.Replace("ControlType.", "");
        return TranslateControlType(typeName);
    }

    private static string TranslateControlType(string controlType)
    {
        var translations = new Dictionary<string, string>
        {
            { "Button", "Przycisk" },
            { "CheckBox", "Pole wyboru" },
            { "ComboBox", "Lista rozwijana" },
            { "Edit", "Pole edycji" },
            { "Hyperlink", "Łącze" },
            { "Image", "Obraz" },
            { "ListItem", "Element listy" },
            { "List", "Lista" },
            { "Menu", "Menu" },
            { "MenuBar", "Pasek menu" },
            { "MenuItem", "Element menu" },
            { "ProgressBar", "Pasek postępu" },
            { "RadioButton", "Przycisk opcji" },
            { "ScrollBar", "Pasek przewijania" },
            { "Slider", "Suwak" },
            { "Spinner", "Pole numeryczne" },
            { "StatusBar", "Pasek stanu" },
            { "Tab", "Zakładka" },
            { "TabItem", "Element zakładki" },
            { "Text", "Tekst" },
            { "ToolBar", "Pasek narzędzi" },
            { "ToolTip", "Podpowiedź" },
            { "Tree", "Drzewo" },
            { "TreeItem", "Element drzewa" },
            { "Custom", "Niestandardowy" },
            { "Group", "Grupa" },
            { "Thumb", "Uchwyt" },
            { "DataGrid", "Tabela danych" },
            { "DataItem", "Element danych" },
            { "Document", "Dokument" },
            { "SplitButton", "Przycisk podzielony" },
            { "Window", "Okno" },
            { "Pane", "Panel" },
            { "Header", "Nagłówek" },
            { "HeaderItem", "Element nagłówka" },
            { "Table", "Tabela" },
            { "TitleBar", "Pasek tytułu" },
            { "Separator", "Separator" }
        };

        return translations.TryGetValue(controlType, out var polish) ? polish : controlType;
    }

    private static string GetElementValue(AutomationElement element)
    {
        try
        {
            // Try ValuePattern
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object? valuePattern))
            {
                var value = ((ValuePattern)valuePattern).Current.Value;
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            // Try TextPattern
            if (element.TryGetCurrentPattern(TextPattern.Pattern, out object? textPattern))
            {
                var text = ((TextPattern)textPattern).DocumentRange.GetText(-1);
                if (!string.IsNullOrWhiteSpace(text))
                    return text.Length > 100 ? text.Substring(0, 100) + "..." : text;
            }

            // Try RangeValuePattern (for sliders, scrollbars, etc.)
            if (element.TryGetCurrentPattern(RangeValuePattern.Pattern, out object? rangePattern))
            {
                var rangeValue = ((RangeValuePattern)rangePattern).Current.Value;
                return rangeValue.ToString();
            }

            // Try TogglePattern (for checkboxes, toggle buttons)
            if (element.TryGetCurrentPattern(TogglePattern.Pattern, out object? togglePattern))
            {
                var toggleState = ((TogglePattern)togglePattern).Current.ToggleState;
                return toggleState.ToString();
            }

            // Try SelectionItemPattern
            if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object? selectionPattern))
            {
                var isSelected = ((SelectionItemPattern)selectionPattern).Current.IsSelected;
                return isSelected ? "Selected" : "Not selected";
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting element value: {ex.Message}");
        }

        return string.Empty;
    }

    public static AutomationElement? GetNextSibling(AutomationElement? element)
    {
        if (element == null)
            return null;

        try
        {
            var walker = TreeWalker.ControlViewWalker;
            return walker.GetNextSibling(element);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting next sibling: {ex.Message}");
            return null;
        }
    }

    public static AutomationElement? GetPreviousSibling(AutomationElement? element)
    {
        if (element == null)
            return null;

        try
        {
            var walker = TreeWalker.ControlViewWalker;
            return walker.GetPreviousSibling(element);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting previous sibling: {ex.Message}");
            return null;
        }
    }

    public static AutomationElement? GetParent(AutomationElement? element)
    {
        if (element == null)
            return null;

        try
        {
            var walker = TreeWalker.ControlViewWalker;
            return walker.GetParent(element);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting parent: {ex.Message}");
            return null;
        }
    }

    public static AutomationElement? GetFirstChild(AutomationElement? element)
    {
        if (element == null)
            return null;

        try
        {
            var walker = TreeWalker.ControlViewWalker;
            return walker.GetFirstChild(element);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting first child: {ex.Message}");
            return null;
        }
    }

    public static bool IsListItem(AutomationElement? element)
    {
        if (element == null)
            return false;

        try
        {
            return element.Current.ControlType == ControlType.ListItem ||
                   element.Current.ControlType == ControlType.DataItem ||
                   element.Current.ControlType == ControlType.TreeItem;
        }
        catch
        {
            return false;
        }
    }

    public static bool IsWindow(AutomationElement? element)
    {
        if (element == null)
            return false;

        try
        {
            return element.Current.ControlType == ControlType.Window;
        }
        catch
        {
            return false;
        }
    }

    public static bool IsButton(AutomationElement? element)
    {
        if (element == null)
            return false;

        try
        {
            return element.Current.ControlType == ControlType.Button;
        }
        catch
        {
            return false;
        }
    }

    public static float GetListItemPosition(AutomationElement? element)
    {
        if (element == null || !IsListItem(element))
            return 0.5f;

        try
        {
            var parent = GetParent(element);
            if (parent == null)
                return 0.5f;

            // Count siblings
            int totalItems = 0;
            int currentIndex = 0;
            bool foundCurrent = false;

            var walker = TreeWalker.ControlViewWalker;
            var firstChild = walker.GetFirstChild(parent);
            var sibling = firstChild;

            while (sibling != null)
            {
                if (IsListItem(sibling))
                {
                    // Add null check before Automation.Compare
                    if (!foundCurrent && element != null && sibling != null)
                    {
                        try
                        {
                            if (Automation.Compare(sibling, element))
                            {
                                currentIndex = totalItems;
                                foundCurrent = true;
                            }
                        }
                        catch
                        {
                            // Ignore compare errors
                        }
                    }
                    totalItems++;
                }
                sibling = walker.GetNextSibling(sibling);
            }

            if (totalItems <= 1)
                return 0.5f;

            return (float)currentIndex / (totalItems - 1);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd wykrywania pozycji: {ex.Message}");
            return 0.5f;
        }
    }

    public static bool IsAtEdge(AutomationElement? element, bool checkNext)
    {
        if (element == null)
            return false;

        try
        {
            var sibling = checkNext ? GetNextSibling(element) : GetPreviousSibling(element);
            return sibling == null;
        }
        catch
        {
            return false;
        }
    }
}

