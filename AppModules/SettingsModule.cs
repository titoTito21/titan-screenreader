using System.Windows.Automation;

namespace ScreenReader.AppModules;

/// <summary>
/// Moduł dla aplikacji Ustawienia Windows
/// Zapewnia ulepszone odczytywanie kategorii i opcji ustawień
/// </summary>
public class SettingsModule : UWPModule
{
    private string? _lastCategory;
    private string? _lastSetting;

    public SettingsModule() : base("SystemSettings")
    {
    }

    public override void OnGainFocus(AutomationElement element)
    {
        base.OnGainFocus(element);
        Console.WriteLine("SettingsModule: Ustawienia aktywne");
    }

    public override void OnFocusChanged(AutomationElement element)
    {
        base.OnFocusChanged(element);

        try
        {
            var controlType = element.Current.ControlType;

            if (controlType == ControlType.ListItem)
            {
                AnnounceSettingItem(element);
            }
            else if (controlType == ControlType.TreeItem)
            {
                AnnounceCategory(element);
            }
            else if (IsToggleSwitch(element))
            {
                AnnounceToggle(element);
            }
            else if (controlType == ControlType.ComboBox)
            {
                AnnounceComboBox(element);
            }
            else if (controlType == ControlType.Slider)
            {
                AnnounceSlider(element);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"SettingsModule: Błąd: {ex.Message}");
        }
    }

    /// <summary>
    /// Sprawdza czy element jest przełącznikiem
    /// </summary>
    private bool IsToggleSwitch(AutomationElement element)
    {
        try
        {
            var controlType = element.Current.ControlType;
            var className = element.Current.ClassName;

            return controlType == ControlType.Button &&
                   (className.Contains("ToggleSwitch") ||
                    element.Current.AutomationId.Contains("Toggle"));
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Ogłasza element ustawienia
    /// </summary>
    private void AnnounceSettingItem(AutomationElement element)
    {
        try
        {
            string name = element.Current.Name;
            string description = GetSettingDescription(element);

            if (!string.IsNullOrEmpty(name) && name != _lastSetting)
            {
                _lastSetting = name;

                if (!string.IsNullOrEmpty(description))
                {
                    Console.WriteLine($"SettingsModule: {name} - {description}");
                }
                else
                {
                    Console.WriteLine($"SettingsModule: {name}");
                }
            }
        }
        catch
        {
            // Ignore
        }
    }

    /// <summary>
    /// Ogłasza kategorię ustawień
    /// </summary>
    private void AnnounceCategory(AutomationElement element)
    {
        try
        {
            string name = element.Current.Name;

            if (!string.IsNullOrEmpty(name) && name != _lastCategory)
            {
                _lastCategory = name;
                Console.WriteLine($"SettingsModule: Kategoria {name}");
            }
        }
        catch
        {
            // Ignore
        }
    }

    /// <summary>
    /// Ogłasza przełącznik
    /// </summary>
    private void AnnounceToggle(AutomationElement element)
    {
        try
        {
            string name = element.Current.Name;
            bool isOn = false;

            // Sprawdź stan przełącznika
            if (element.TryGetCurrentPattern(TogglePattern.Pattern, out var pattern))
            {
                var togglePattern = (TogglePattern)pattern;
                isOn = togglePattern.Current.ToggleState == ToggleState.On;
            }

            string state = isOn ? "włączony" : "wyłączony";
            Console.WriteLine($"SettingsModule: Przełącznik {name}, {state}");
        }
        catch
        {
            // Ignore
        }
    }

    /// <summary>
    /// Ogłasza pole wyboru
    /// </summary>
    private void AnnounceComboBox(AutomationElement element)
    {
        try
        {
            string name = element.Current.Name;
            string value = "";

            // Pobierz aktualną wartość
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var pattern))
            {
                var valuePattern = (ValuePattern)pattern;
                value = valuePattern.Current.Value;
            }
            else if (element.TryGetCurrentPattern(SelectionPattern.Pattern, out var selPattern))
            {
                var selectionPattern = (SelectionPattern)selPattern;
                var selection = selectionPattern.Current.GetSelection();
                if (selection.Length > 0)
                {
                    value = selection[0].Current.Name;
                }
            }

            if (!string.IsNullOrEmpty(value))
            {
                Console.WriteLine($"SettingsModule: Lista {name}, wybrano: {value}");
            }
            else
            {
                Console.WriteLine($"SettingsModule: Lista {name}");
            }
        }
        catch
        {
            // Ignore
        }
    }

    /// <summary>
    /// Ogłasza suwak
    /// </summary>
    private void AnnounceSlider(AutomationElement element)
    {
        try
        {
            string name = element.Current.Name;
            double value = 0;

            if (element.TryGetCurrentPattern(RangeValuePattern.Pattern, out var pattern))
            {
                var rangePattern = (RangeValuePattern)pattern;
                value = rangePattern.Current.Value;
                double min = rangePattern.Current.Minimum;
                double max = rangePattern.Current.Maximum;

                // Oblicz procent
                double percent = (value - min) / (max - min) * 100;

                Console.WriteLine($"SettingsModule: Suwak {name}, {percent:F0}%");
            }
            else
            {
                Console.WriteLine($"SettingsModule: Suwak {name}");
            }
        }
        catch
        {
            // Ignore
        }
    }

    /// <summary>
    /// Pobiera opis ustawienia
    /// </summary>
    private string GetSettingDescription(AutomationElement element)
    {
        try
        {
            // Próbuj pobrać HelpText
            string helpText = element.Current.HelpText;
            if (!string.IsNullOrEmpty(helpText))
                return helpText;

            // Próbuj znaleźć opis w dzieciach
            var walker = TreeWalker.ControlViewWalker;
            var child = walker.GetFirstChild(element);
            while (child != null)
            {
                var childType = child.Current.ControlType;
                if (childType == ControlType.Text)
                {
                    var text = child.Current.Name;
                    if (!string.IsNullOrEmpty(text) && text != element.Current.Name)
                    {
                        return text;
                    }
                }
                child = walker.GetNextSibling(child);
            }
        }
        catch
        {
            // Ignore
        }

        return "";
    }

    public override void CustomizeElement(AutomationElement element, ref string name, ref string role)
    {
        base.CustomizeElement(element, ref name, ref role);

        try
        {
            var controlType = element.Current.ControlType;

            if (controlType == ControlType.ListItem)
            {
                // Dodaj opis jeśli dostępny
                string description = GetSettingDescription(element);
                if (!string.IsNullOrEmpty(description) && !name.Contains(description))
                {
                    name = $"{name}, {description}";
                }
                role = "ustawienie";
            }
            else if (controlType == ControlType.TreeItem)
            {
                role = "kategoria";
            }
            else if (IsToggleSwitch(element))
            {
                role = "przełącznik";
            }
        }
        catch
        {
            // Ignore
        }
    }

    public override bool ShouldUseVirtualBuffer(AutomationElement element)
    {
        // W Ustawieniach używamy wirtualnego bufora dla list i dokumentów
        try
        {
            var controlType = element.Current.ControlType;
            return controlType == ControlType.List ||
                   controlType == ControlType.Document;
        }
        catch
        {
            return false;
        }
    }
}
