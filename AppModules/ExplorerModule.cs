using System.Windows.Automation;

namespace ScreenReader.AppModules;

/// <summary>
/// Moduł dla Windows Explorer
/// Port z NVDA appModules/explorer.py - zaawansowana wersja
/// </summary>
public class ExplorerModule : AppModuleBase
{
    // Typy kontrolek specyficzne dla Explorera
    private static readonly string[] TreeItemClasses = { "SysTreeView32", "TreeView" };
    private static readonly string[] ListViewClasses = { "SysListView32", "DirectUIHWND" };

    public ExplorerModule() : base("explorer")
    {
    }

    public override void OnGainFocus(AutomationElement element)
    {
        base.OnGainFocus(element);
        Console.WriteLine("ExplorerModule: Fokus na eksploratorze");
    }

    public override void OnFocusChanged(AutomationElement element)
    {
        var controlType = element.Current.ControlType;

        // Drzewo folderów (lewy panel)
        if (controlType == ControlType.TreeItem)
        {
            AnnounceTreeItem(element);
        }
        // Lista plików
        else if (controlType == ControlType.ListItem || controlType == ControlType.DataItem)
        {
            AnnounceFileItem(element);
        }
    }

    public override void CustomizeElement(AutomationElement element, ref string name, ref string role)
    {
        var controlType = element.Current.ControlType;

        if (controlType == ControlType.ListItem || controlType == ControlType.DataItem)
        {
            // Dodaj informacje o pliku
            var fileInfo = GetFileItemInfo(element);
            if (!string.IsNullOrEmpty(fileInfo))
            {
                name = $"{name}, {fileInfo}";
            }
        }
    }

    /// <summary>
    /// Ogłasza element drzewa folderów
    /// </summary>
    private void AnnounceTreeItem(AutomationElement element)
    {
        try
        {
            string name = element.Current.Name;
            string expandState = GetExpandCollapseState(element);
            int level = GetTreeItemLevel(element);

            string announcement = name;
            if (!string.IsNullOrEmpty(expandState))
                announcement += $", {expandState}";
            if (level > 0)
                announcement += $", poziom {level}";

            Console.WriteLine($"ExplorerModule: Folder {announcement}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ExplorerModule: Błąd przy ogłaszaniu drzewa: {ex.Message}");
        }
    }

    /// <summary>
    /// Ogłasza element listy plików
    /// </summary>
    private void AnnounceFileItem(AutomationElement element)
    {
        try
        {
            string name = element.Current.Name;
            string fileType = GetFileType(name);
            string size = GetFileSize(element);
            string date = GetFileDate(element);

            var parts = new List<string> { name };

            if (!string.IsNullOrEmpty(fileType))
                parts.Add(fileType);
            if (!string.IsNullOrEmpty(size))
                parts.Add(size);
            if (!string.IsNullOrEmpty(date))
                parts.Add(date);

            string announcement = string.Join(", ", parts);
            Console.WriteLine($"ExplorerModule: Plik {announcement}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ExplorerModule: Błąd przy ogłaszaniu pliku: {ex.Message}");
        }
    }

    /// <summary>
    /// Pobiera stan rozwinięcia/zwinięcia elementu
    /// </summary>
    private string GetExpandCollapseState(AutomationElement element)
    {
        try
        {
            if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var pattern))
            {
                var expandPattern = (ExpandCollapsePattern)pattern;
                return expandPattern.Current.ExpandCollapseState switch
                {
                    ExpandCollapseState.Expanded => "rozwinięty",
                    ExpandCollapseState.Collapsed => "zwinięty",
                    ExpandCollapseState.PartiallyExpanded => "częściowo rozwinięty",
                    _ => ""
                };
            }
        }
        catch { }
        return "";
    }

    /// <summary>
    /// Pobiera poziom zagnieżdżenia w drzewie
    /// </summary>
    private int GetTreeItemLevel(AutomationElement element)
    {
        int level = 0;
        var current = element;

        while (current != null)
        {
            var parent = TreeWalker.ControlViewWalker.GetParent(current);
            if (parent == null || parent.Current.ControlType == ControlType.Tree)
                break;

            if (parent.Current.ControlType == ControlType.TreeItem)
                level++;

            current = parent;
        }

        return level;
    }

    /// <summary>
    /// Określa typ pliku na podstawie nazwy
    /// </summary>
    private string GetFileType(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            return "";

        // Sprawdź czy to folder (brak rozszerzenia lub specjalne nazwy)
        if (!fileName.Contains('.'))
            return "Folder";

        string ext = Path.GetExtension(fileName).ToLowerInvariant();

        return ext switch
        {
            ".txt" => "dokument tekstowy",
            ".doc" or ".docx" => "dokument Word",
            ".xls" or ".xlsx" => "arkusz Excel",
            ".ppt" or ".pptx" => "prezentacja PowerPoint",
            ".pdf" => "dokument PDF",
            ".jpg" or ".jpeg" or ".png" or ".gif" or ".bmp" => "obraz",
            ".mp3" or ".wav" or ".flac" or ".ogg" => "plik audio",
            ".mp4" or ".avi" or ".mkv" or ".mov" => "plik wideo",
            ".zip" or ".rar" or ".7z" => "archiwum",
            ".exe" => "aplikacja",
            ".dll" => "biblioteka DLL",
            ".cs" => "plik C#",
            ".py" => "plik Python",
            ".js" => "plik JavaScript",
            ".html" or ".htm" => "strona HTML",
            ".css" => "arkusz stylów",
            ".json" => "plik JSON",
            ".xml" => "plik XML",
            ".md" => "plik Markdown",
            _ => $"plik {ext.TrimStart('.')}"
        };
    }

    /// <summary>
    /// Pobiera rozmiar pliku z właściwości elementu
    /// </summary>
    private string GetFileSize(AutomationElement element)
    {
        try
        {
            // Próbuj odczytać z GridItemPattern lub z dzieci elementu
            // Explorer w widoku szczegółów używa kolumn jako dzieci

            var walker = TreeWalker.ControlViewWalker;
            var child = walker.GetFirstChild(element);
            int columnIndex = 0;

            while (child != null)
            {
                // Kolumna 1 = Rozmiar (w typowym widoku szczegółów)
                // Kolejność: Nazwa, Data modyfikacji, Typ, Rozmiar
                // Ale może się różnić w zależności od ustawień
                try
                {
                    var childName = child.Current.Name;
                    var childControlType = child.Current.ControlType;

                    // Sprawdź czy to wygląda jak rozmiar (zawiera KB, MB, GB, bajty)
                    if (!string.IsNullOrEmpty(childName))
                    {
                        if (childName.Contains(" KB") || childName.Contains(" MB") ||
                            childName.Contains(" GB") || childName.Contains(" bajt"))
                        {
                            return childName;
                        }
                    }
                }
                catch
                {
                    // Kontynuuj z następnym dzieckiem
                }

                child = walker.GetNextSibling(child);
                columnIndex++;

                // Bezpieczeństwo - nie szukaj w nieskończoność
                if (columnIndex > 10)
                    break;
            }
        }
        catch { }
        return "";
    }

    /// <summary>
    /// Pobiera datę modyfikacji pliku
    /// </summary>
    private string GetFileDate(AutomationElement element)
    {
        try
        {
            var walker = TreeWalker.ControlViewWalker;
            var child = walker.GetFirstChild(element);
            int columnIndex = 0;

            while (child != null)
            {
                try
                {
                    var childName = child.Current.Name;

                    // Sprawdź czy to wygląda jak data (zawiera rok lub format daty)
                    if (!string.IsNullOrEmpty(childName))
                    {
                        // Wzorce dat
                        if (System.Text.RegularExpressions.Regex.IsMatch(childName,
                            @"\d{1,4}[./-]\d{1,2}[./-]\d{1,4}"))
                        {
                            return childName;
                        }
                    }
                }
                catch { }

                child = walker.GetNextSibling(child);
                columnIndex++;

                if (columnIndex > 10)
                    break;
            }
        }
        catch { }
        return "";
    }

    /// <summary>
    /// Pobiera dodatkowe informacje o pliku
    /// </summary>
    private string GetFileItemInfo(AutomationElement element)
    {
        var parts = new List<string>();

        string size = GetFileSize(element);
        if (!string.IsNullOrEmpty(size))
            parts.Add(size);

        string date = GetFileDate(element);
        if (!string.IsNullOrEmpty(date))
            parts.Add(date);

        return string.Join(", ", parts);
    }

    /// <summary>
    /// Obsługuje okno właściwości pliku
    /// </summary>
    public void HandlePropertyWindow(AutomationElement window)
    {
        try
        {
            // Znajdź aktywną zakładkę
            var tabControl = window.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tab));

            if (tabControl != null)
            {
                var selectedTab = tabControl.FindFirst(TreeScope.Children,
                    new PropertyCondition(SelectionItemPattern.IsSelectedProperty, true));

                if (selectedTab != null)
                {
                    Console.WriteLine($"ExplorerModule: Właściwości, zakładka {selectedTab.Current.Name}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ExplorerModule: Błąd przy obsłudze właściwości: {ex.Message}");
        }
    }

    /// <summary>
    /// Obsługuje panel podglądu
    /// </summary>
    public string? GetPreviewPaneContent(AutomationElement pane)
    {
        try
        {
            var document = pane.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document));

            if (document != null)
            {
                if (document.TryGetCurrentPattern(TextPattern.Pattern, out var pattern))
                {
                    var textPattern = (TextPattern)pattern;
                    return textPattern.DocumentRange.GetText(500);
                }
            }
        }
        catch { }
        return null;
    }

    public override string? GetStatusBarText()
    {
        try
        {
            // Znajdź okno eksploratora
            var explorerWindow = FindExplorerWindow();
            if (explorerWindow == null)
                return null;

            // Szukaj paska stanu
            var statusBar = explorerWindow.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.StatusBar));

            if (statusBar != null)
            {
                // Pobierz tekst z paska stanu
                var text = statusBar.Current.Name;
                if (!string.IsNullOrEmpty(text))
                    return text;

                // Spróbuj pobrać tekst z dzieci paska stanu
                var textElements = statusBar.FindAll(TreeScope.Children,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));

                var parts = new List<string>();
                foreach (AutomationElement textEl in textElements)
                {
                    try
                    {
                        var name = textEl.Current.Name;
                        if (!string.IsNullOrEmpty(name))
                            parts.Add(name);
                    }
                    catch { }
                }

                if (parts.Count > 0)
                    return string.Join(", ", parts);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ExplorerModule: Błąd odczytu paska stanu: {ex.Message}");
        }

        return null;
    }

    /// <summary>
    /// Znajduje główne okno eksploratora dla aktualnego procesu
    /// </summary>
    private AutomationElement? FindExplorerWindow()
    {
        try
        {
            if (ProcessId == 0)
                return null;

            var condition = new AndCondition(
                new PropertyCondition(AutomationElement.ProcessIdProperty, ProcessId),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

            return AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Pobiera liczbę zaznaczonych elementów
    /// </summary>
    public int GetSelectedItemCount()
    {
        try
        {
            var window = FindExplorerWindow();
            if (window == null)
                return 0;

            // Znajdź listę plików
            var listView = window.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List));

            if (listView != null && listView.TryGetCurrentPattern(SelectionPattern.Pattern, out var pattern))
            {
                var selectionPattern = (SelectionPattern)pattern;
                var selection = selectionPattern.Current.GetSelection();
                return selection.Length;
            }
        }
        catch { }

        return 0;
    }

    /// <summary>
    /// Ogłasza liczbę zaznaczonych elementów
    /// </summary>
    public string GetSelectionAnnouncement()
    {
        int count = GetSelectedItemCount();
        if (count == 0)
            return "Brak zaznaczonych elementów";
        else if (count == 1)
            return "Zaznaczono 1 element";
        else if (count >= 2 && count <= 4)
            return $"Zaznaczono {count} elementy";
        else
            return $"Zaznaczono {count} elementów";
    }
}
