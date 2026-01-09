using System.Windows.Automation;

namespace ScreenReader.UIAutomation;

/// <summary>
/// Detektor elementów UI z zaawansowanym skanowaniem rekursywnym
/// Znajduje nawet małe elementy (ikony pulpitu, list items) w promieniu punktu
/// </summary>
public static class ElementDetector
{
    /// <summary>
    /// Znajduje element UI w punkcie z rekursywnym skanowaniem drzewa
    /// </summary>
    /// <param name="x">Współrzędna X ekranu</param>
    /// <param name="y">Współrzędna Y ekranu</param>
    /// <param name="searchRadius">Promień wyszukiwania w pikselach (domyślnie 10px)</param>
    /// <returns>Znaleziony element lub null</returns>
    public static AutomationElement? FindElementAtPoint(int x, int y, int searchRadius = 10)
    {
        var point = new System.Windows.Point(x, y);

        try
        {
            // Krok 1: Spróbuj bezpośredniego trafienia
            var element = AutomationElement.FromPoint(point);

            if (element != null && IsInteractiveElement(element))
            {
                return element;
            }

            // Krok 2: Rekursywne skanowanie drzewa w promieniu
            var candidates = new List<(AutomationElement elem, double distance)>();

            if (element != null)
            {
                ScanTreeInRadius(element, x, y, searchRadius, candidates);
            }

            // Krok 3: Zwróć najbliższy interaktywny element
            if (candidates.Count > 0)
            {
                var closest = candidates.OrderBy(c => c.distance).First();
                Console.WriteLine($"ElementDetector: Znaleziono {candidates.Count} kandydatów, wybrany: {GetElementDescription(closest.elem)} (distance={closest.distance:F1}px)");
                return closest.elem;
            }

            // Krok 4: Fallback - zwróć element z FromPoint() nawet jeśli nieinteraktywny
            return element;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ElementDetector: Błąd wykrywania: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Skanuje drzewo UI rekursywnie w promieniu punktu
    /// </summary>
    private static void ScanTreeInRadius(
        AutomationElement root,
        int x,
        int y,
        int radius,
        List<(AutomationElement, double)> candidates,
        int depth = 0)
    {
        // Limit głębokości rekursji aby uniknąć zbyt długiego czasu
        if (depth > 5 || root == null)
            return;

        var walker = TreeWalker.ControlViewWalker;
        AutomationElement? child = null;

        try
        {
            child = walker.GetFirstChild(root);
        }
        catch (ElementNotAvailableException)
        {
            return;
        }

        while (child != null)
        {
            try
            {
                var rect = child.Current.BoundingRectangle;

                if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
                {
                    // Oblicz odległość od punktu do centrum elementu
                    double centerX = rect.X + rect.Width / 2;
                    double centerY = rect.Y + rect.Height / 2;
                    double distance = Math.Sqrt(Math.Pow(centerX - x, 2) + Math.Pow(centerY - y, 2));

                    // Jeśli element jest w promieniu i interaktywny, dodaj do kandydatów
                    if (distance <= radius && IsInteractiveElement(child))
                    {
                        candidates.Add((child, distance));
                    }

                    // Rekursywnie skanuj dzieci
                    ScanTreeInRadius(child, x, y, radius, candidates, depth + 1);
                }
            }
            catch (ElementNotAvailableException)
            {
                // Element już nieaktualny, kontynuuj z następnym
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ElementDetector: Błąd skanowania dziecka: {ex.Message}");
            }

            // Przejdź do następnego rodzeństwa
            try
            {
                child = walker.GetNextSibling(child);
            }
            catch (ElementNotAvailableException)
            {
                break;
            }
        }
    }

    /// <summary>
    /// Sprawdza czy element jest interaktywny (warto go zapowiedzieć)
    /// </summary>
    private static bool IsInteractiveElement(AutomationElement element)
    {
        try
        {
            var controlType = element.Current.ControlType;
            var isKeyboardFocusable = element.Current.IsKeyboardFocusable;

            // Typy kontrolek uważane za interaktywne
            bool isInteractiveType =
                controlType == ControlType.Button ||
                controlType == ControlType.Edit ||
                controlType == ControlType.Hyperlink ||
                controlType == ControlType.ListItem ||
                controlType == ControlType.MenuItem ||
                controlType == ControlType.CheckBox ||
                controlType == ControlType.RadioButton ||
                controlType == ControlType.ComboBox ||
                controlType == ControlType.Slider ||
                controlType == ControlType.TabItem ||
                controlType == ControlType.TreeItem ||
                controlType == ControlType.DataItem ||
                controlType == ControlType.SplitButton ||
                controlType == ControlType.Custom; // Custom może być ikoną pulpitu

            // Element jest interaktywny jeśli:
            // - Ma interaktywny typ ORAZ ma nazwę (nie pusty)
            // - LUB jest focusable z klawiatury
            bool hasName = !string.IsNullOrWhiteSpace(element.Current.Name);

            return (isInteractiveType && hasName) || isKeyboardFocusable;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Zwraca krótki opis elementu dla debugowania
    /// </summary>
    private static string GetElementDescription(AutomationElement element)
    {
        try
        {
            string name = element.Current.Name;
            string type = element.Current.ControlType.ProgrammaticName.Replace("ControlType.", "");

            if (string.IsNullOrWhiteSpace(name))
                name = "(bez nazwy)";

            return $"{type}: {name}";
        }
        catch
        {
            return "(nieznany)";
        }
    }
}
