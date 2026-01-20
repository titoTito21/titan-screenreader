using System.Windows.Automation;
using ScreenReader.Settings;

namespace ScreenReader;

/// <summary>
/// Informacje o elemencie UI w rozbiciu na części
/// </summary>
public class ElementInfo
{
    public string Name { get; set; } = "";
    public string ControlType { get; set; } = "";
    public string ControlTypePolish { get; set; } = "";
    public string Value { get; set; } = "";
    public string State { get; set; } = "";
    public string HelpText { get; set; } = "";
    public string PositionInfo { get; set; } = "";
    public bool IsBasicControl { get; set; }
    public bool IsBlockControl { get; set; }
}

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

    /// <summary>
    /// Pobiera informacje o elemencie w rozbiciu na części
    /// </summary>
    public static ElementInfo GetElementInfo(AutomationElement? element)
    {
        var info = new ElementInfo();

        if (element == null)
            return info;

        try
        {
            var controlType = element.Current.ControlType;
            var controlTypeName = controlType.ProgrammaticName.Replace("ControlType.", "");

            info.Name = element.Current.Name ?? "";
            info.ControlType = controlTypeName;
            info.ControlTypePolish = TranslateControlType(controlTypeName);
            info.Value = GetElementValue(element);
            info.HelpText = element.Current.HelpText ?? "";

            // Określ kategorię kontrolki
            info.IsBasicControl = IsBasicControlType(controlTypeName);
            info.IsBlockControl = IsBlockControlType(controlTypeName);

            // Pobierz pozycję dla elementów listy
            if (controlType == ControlType.ListItem ||
                controlType == ControlType.DataItem ||
                controlType == ControlType.TreeItem)
            {
                info.PositionInfo = GetListItemPositionInfo(element);
            }

            // Pobierz stan dla określonych kontrolek
            info.State = GetElementState(element);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd GetElementInfo: {ex.Message}");
        }

        return info;
    }

    /// <summary>
    /// Sprawdza czy to podstawowa kontrolka (przycisk, pole edycji, pole wyboru)
    /// </summary>
    private static bool IsBasicControlType(string controlType)
    {
        return controlType switch
        {
            "Button" or "Edit" or "CheckBox" or "RadioButton" or "ComboBox" or
            "Slider" or "Spinner" or "Hyperlink" or "SplitButton" => true,
            _ => false
        };
    }

    /// <summary>
    /// Sprawdza czy to kontrolka blokowa (element listy, element menu, etc)
    /// </summary>
    private static bool IsBlockControlType(string controlType)
    {
        return controlType switch
        {
            "ListItem" or "DataItem" or "TreeItem" or "MenuItem" or "TabItem" or "HeaderItem" => true,
            _ => false
        };
    }

    /// <summary>
    /// Pobiera stan elementu (zaznaczony, rozwinięty, itp.)
    /// </summary>
    private static string GetElementState(AutomationElement element)
    {
        var states = new List<string>();

        try
        {
            // TogglePattern - pola wyboru
            if (element.TryGetCurrentPattern(TogglePattern.Pattern, out object? togglePattern))
            {
                var toggleState = ((TogglePattern)togglePattern).Current.ToggleState;
                states.Add(toggleState switch
                {
                    ToggleState.On => "zaznaczono",
                    ToggleState.Off => "odznaczono",
                    ToggleState.Indeterminate => "częściowo zaznaczono",
                    _ => ""
                });
            }

            // SelectionItemPattern - elementy selekcji
            if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object? selectionPattern))
            {
                var isSelected = ((SelectionItemPattern)selectionPattern).Current.IsSelected;
                if (isSelected)
                    states.Add("zaznaczony");
            }

            // ExpandCollapsePattern - elementy rozwijane
            if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out object? expandPattern))
            {
                var expandState = ((ExpandCollapsePattern)expandPattern).Current.ExpandCollapseState;
                states.Add(expandState switch
                {
                    ExpandCollapseState.Expanded => "rozwinięto",
                    ExpandCollapseState.Collapsed => "zwinięto",
                    ExpandCollapseState.PartiallyExpanded => "częściowo rozwinięto",
                    _ => ""
                });
            }

            // Sprawdź IsEnabled
            if (!element.Current.IsEnabled)
            {
                states.Add("niedostępny");
            }

            // Sprawdź IsOffscreen
            if (element.Current.IsOffscreen)
            {
                states.Add("poza ekranem");
            }
        }
        catch { }

        return string.Join(", ", states.Where(s => !string.IsNullOrEmpty(s)));
    }

    /// <summary>
    /// Formatuje opis elementu zgodnie z ustawieniami szczegółowości
    /// </summary>
    public static string FormatElementDescription(ElementInfo info, SettingsManager settings)
    {
        var parts = new List<string>();

        // Nazwa elementu
        if (settings.ElementName && !string.IsNullOrWhiteSpace(info.Name))
        {
            parts.Add(info.Name);
        }

        // Typ kontrolki
        if (settings.ElementType)
        {
            bool shouldAnnounceType = true;

            // Sprawdź ustawienia dla podstawowych/blokowych kontrolek
            if (info.IsBasicControl && !settings.AnnounceBasicControls)
            {
                shouldAnnounceType = false;
            }
            else if (info.IsBlockControl && !settings.AnnounceBlockControls)
            {
                shouldAnnounceType = false;
            }

            if (shouldAnnounceType && !string.IsNullOrWhiteSpace(info.ControlTypePolish))
            {
                parts.Add(info.ControlTypePolish);
            }
        }

        // Pozycja na liście
        if (settings.AnnounceListPosition && !string.IsNullOrWhiteSpace(info.PositionInfo))
        {
            parts.Add(info.PositionInfo);
        }

        // Stan kontrolki
        if (settings.ElementState && !string.IsNullOrWhiteSpace(info.State))
        {
            parts.Add(info.State);
        }

        // Wartość (parametr kontrolki)
        if (settings.ElementParameter && !string.IsNullOrWhiteSpace(info.Value))
        {
            // Nie dodawaj wartości jeśli jest taka sama jak stan
            if (info.Value != info.State)
            {
                parts.Add(info.Value);
            }
        }

        // Help text
        if (!string.IsNullOrWhiteSpace(info.HelpText))
        {
            parts.Add(info.HelpText);
        }

        return string.Join(", ", parts);
    }

    public static string GetElementDescription(AutomationElement? element)
    {
        if (element == null)
            return "Brak elementu";

        try
        {
            var name = element.Current.Name;
            var controlType = element.Current.ControlType;
            var controlTypeName = controlType.ProgrammaticName.Replace("ControlType.", "");
            var controlTypePolish = TranslateControlType(controlTypeName);
            var value = GetElementValue(element);
            var helpText = element.Current.HelpText;

            var description = string.IsNullOrWhiteSpace(name) ? controlTypePolish : $"{name}, {controlTypePolish}";

            // Dla elementów listy dodaj pozycję "X z Z"
            if (controlType == ControlType.ListItem ||
                controlType == ControlType.DataItem ||
                controlType == ControlType.TreeItem)
            {
                var positionInfo = GetListItemPositionInfo(element);
                if (!string.IsNullOrEmpty(positionInfo))
                {
                    description += $", {positionInfo}";
                }
            }

            // Dla elementów listy/drzewa dodaj informację o rodzicu (lista/drzewo)
            if (controlType == ControlType.ListItem ||
                controlType == ControlType.DataItem ||
                controlType == ControlType.TreeItem)
            {
                var parentInfo = GetParentContainerInfo(element);
                if (!string.IsNullOrEmpty(parentInfo))
                {
                    description = $"{parentInfo}, {description}";
                }
            }

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
    /// Pobiera informacje o pozycji elementu na liście "X z Z"
    /// </summary>
    public static string GetListItemPositionInfo(AutomationElement? element)
    {
        if (element == null)
            return "";

        try
        {
            var parent = GetParent(element);
            if (parent == null)
                return "";

            int totalItems = 0;
            int currentIndex = 0;
            bool foundCurrent = false;

            var walker = TreeWalker.ControlViewWalker;
            var sibling = walker.GetFirstChild(parent);

            while (sibling != null)
            {
                var siblingType = sibling.Current.ControlType;
                if (siblingType == ControlType.ListItem ||
                    siblingType == ControlType.DataItem ||
                    siblingType == ControlType.TreeItem)
                {
                    totalItems++;
                    if (!foundCurrent)
                    {
                        try
                        {
                            if (Automation.Compare(sibling, element))
                            {
                                currentIndex = totalItems;
                                foundCurrent = true;
                            }
                        }
                        catch { }
                    }
                }
                sibling = walker.GetNextSibling(sibling);
            }

            if (totalItems > 0 && currentIndex > 0)
            {
                return $"{currentIndex} z {totalItems}";
            }
        }
        catch { }

        return "";
    }

    /// <summary>
    /// Pobiera informację o kontenerze rodzica (np. "Lista", "Drzewo")
    /// </summary>
    public static string GetParentContainerInfo(AutomationElement? element)
    {
        if (element == null)
            return "";

        try
        {
            var parent = GetParent(element);
            if (parent == null)
                return "";

            var parentType = parent.Current.ControlType;
            if (parentType == ControlType.List)
                return "Lista";
            if (parentType == ControlType.Tree)
                return "Drzewo";
            if (parentType == ControlType.DataGrid)
                return "Tabela danych";
            if (parentType == ControlType.Menu)
                return "Menu";
            if (parentType == ControlType.ComboBox)
                return "Lista rozwijana";
        }
        catch { }

        return "";
    }

    /// <summary>
    /// Pobiera element listy będący rodzicem elementu
    /// </summary>
    public static AutomationElement? GetListParent(AutomationElement? element)
    {
        if (element == null)
            return null;

        try
        {
            var parent = GetParent(element);
            if (parent == null)
                return null;

            var parentType = parent.Current.ControlType;
            if (parentType == ControlType.List ||
                parentType == ControlType.Tree ||
                parentType == ControlType.DataGrid ||
                parentType == ControlType.Table)
            {
                return parent;
            }

            // W niektórych przypadkach (np. Eksplorator Windows) lista może być zagnieżdżona
            // Spróbuj poszukać jeszcze jeden poziom wyżej
            var grandParent = GetParent(parent);
            if (grandParent != null)
            {
                var grandParentType = grandParent.Current.ControlType;
                if (grandParentType == ControlType.List ||
                    grandParentType == ControlType.Tree ||
                    grandParentType == ControlType.DataGrid ||
                    grandParentType == ControlType.Table)
                {
                    return grandParent;
                }
            }
        }
        catch { }

        return null;
    }

    /// <summary>
    /// Pobiera etykietę listy (nazwa lub LabeledBy)
    /// </summary>
    public static string GetListLabel(AutomationElement? listElement)
    {
        if (listElement == null)
            return "";

        try
        {
            // Najpierw sprawdź LabeledBy
            var labeledBy = listElement.Current.LabeledBy;
            if (labeledBy != null)
            {
                var labelName = labeledBy.Current.Name;
                if (!string.IsNullOrWhiteSpace(labelName))
                    return labelName;
            }

            // Sprawdź nazwę listy
            var name = listElement.Current.Name;
            if (!string.IsNullOrWhiteSpace(name))
                return name;
        }
        catch { }

        return "";
    }

    /// <summary>
    /// Pobiera opis elementu bez informacji o rodzicu lista/drzewo
    /// </summary>
    public static string GetElementDescriptionWithoutListInfo(AutomationElement? element)
    {
        if (element == null)
            return "Brak elementu";

        try
        {
            var name = element.Current.Name;
            var controlType = element.Current.ControlType;
            var controlTypeName = controlType.ProgrammaticName.Replace("ControlType.", "");
            var controlTypePolish = TranslateControlType(controlTypeName);
            var value = GetElementValue(element);
            var helpText = element.Current.HelpText;

            var description = string.IsNullOrWhiteSpace(name) ? controlTypePolish : $"{name}, {controlTypePolish}";

            // Dla elementów listy dodaj pozycję "X z Z"
            if (controlType == ControlType.ListItem ||
                controlType == ControlType.DataItem ||
                controlType == ControlType.TreeItem)
            {
                var positionInfo = GetListItemPositionInfo(element);
                if (!string.IsNullOrEmpty(positionInfo))
                {
                    description += $", {positionInfo}";
                }
            }

            // NIE dodajemy informacji o rodzicu (lista/drzewo) - to robi ScreenReaderEngine

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
                return toggleState switch
                {
                    ToggleState.On => "zaznaczono",
                    ToggleState.Off => "odznaczono",
                    ToggleState.Indeterminate => "częściowo zaznaczono",
                    _ => toggleState.ToString()
                };
            }

            // Try SelectionItemPattern
            if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object? selectionPattern))
            {
                var isSelected = ((SelectionItemPattern)selectionPattern).Current.IsSelected;
                return isSelected ? "zaznaczony" : "niezaznaczony";
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
            var controlType = element.Current.ControlType;

            // Sprawdź standardowe typy elementów listy
            if (controlType == ControlType.ListItem ||
                controlType == ControlType.DataItem ||
                controlType == ControlType.TreeItem)
            {
                return true;
            }

            // Sprawdź czy element ma wzorzec SelectionItem (elementy wybieralne w liście/gridzie)
            // To obejmuje elementy w Eksploratorze Windows, które mogą być typu Custom
            object? pattern;
            if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out pattern))
            {
                // Sprawdź czy rodzic to lista/grid/tree
                var parent = GetParent(element);
                if (parent != null)
                {
                    var parentType = parent.Current.ControlType;
                    if (parentType == ControlType.List ||
                        parentType == ControlType.Tree ||
                        parentType == ControlType.DataGrid ||
                        parentType == ControlType.Table)
                    {
                        return true;
                    }
                }
            }

            return false;
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

