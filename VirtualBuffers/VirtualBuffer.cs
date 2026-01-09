using System.Text;
using System.Windows.Automation;
using ScreenReader.BrowseMode;

namespace ScreenReader.VirtualBuffers;

/// <summary>
/// Właściwości UIA używane przez Edge/Chrome do atrybutów ARIA
/// </summary>
internal static class AriaUiaProperties
{
    // UIA właściwości dla ARIA
    public static readonly AutomationProperty? AriaRoleProperty;
    public static readonly AutomationProperty? AriaPropertiesProperty;
    public static readonly AutomationProperty? LevelProperty;
    public static readonly AutomationProperty? PositionInSetProperty;
    public static readonly AutomationProperty? SizeOfSetProperty;
    public static readonly AutomationProperty? DescribedByProperty;
    public static readonly AutomationProperty? FullDescriptionProperty;

    static AriaUiaProperties()
    {
        try
        {
            // AriaRole - UIA_AriaRolePropertyId = 30101
            AriaRoleProperty = AutomationProperty.LookupById(30101);
            // AriaProperties - UIA_AriaPropertiesPropertyId = 30102
            AriaPropertiesProperty = AutomationProperty.LookupById(30102);
            // Level - UIA_LevelPropertyId = 30154
            LevelProperty = AutomationProperty.LookupById(30154);
            // PositionInSet - UIA_PositionInSetPropertyId = 30152
            PositionInSetProperty = AutomationProperty.LookupById(30152);
            // SizeOfSet - UIA_SizeOfSetPropertyId = 30153
            SizeOfSetProperty = AutomationProperty.LookupById(30153);
            // DescribedBy - UIA_DescribedByPropertyId = 30105
            DescribedByProperty = AutomationProperty.LookupById(30105);
            // FullDescription - UIA_FullDescriptionPropertyId = 30159
            FullDescriptionProperty = AutomationProperty.LookupById(30159);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"VirtualBuffer: Nie można zainicjować właściwości ARIA UIA: {ex.Message}");
        }
    }
}

/// <summary>
/// Wirtualny bufor - płaska reprezentacja tekstowa dokumentu webowego
/// Port z NVDA virtualBuffers/__init__.py
///
/// Zgodnie z koncepcją NVDA:
/// - Przechowuje drzewo UI jako płaski tekst z offsetami
/// - Każdy węzeł ma startOffset i endOffset w buforze tekstowym
/// - Umożliwia nawigację niezależną od fokusa
/// - Używa textInfo API dla zarządzania offsetami
/// - Wspiera ARIA atrybuty przez UIA
/// - Dynamiczne aktualizacje przy zmianach w dokumencie
/// </summary>
public class VirtualBuffer : IDisposable
{
    private readonly StringBuilder _buffer = new();
    private readonly List<VirtualBufferNode> _nodes = new();
    private readonly Dictionary<int, VirtualBufferNode> _nodeCache = new();
    private int _caretOffset;
    private int _nextNodeId;
    private bool _disposed;
    private bool _isRebuilding;

    /// <summary>Dokument źródłowy</summary>
    public AutomationElement? RootElement { get; private set; }

    /// <summary>Czy bufor jest ładowany</summary>
    public bool IsLoading { get; private set; }

    /// <summary>Czy tryb pass-through (focus mode)</summary>
    public bool PassThrough { get; set; }

    /// <summary>Nazwa procesu źródłowego (edge, chrome)</summary>
    public string? ProcessName { get; private set; }

    /// <summary>Czy to przeglądarka Chromium</summary>
    public bool IsChromiumBrowser { get; private set; }

    /// <summary>Aktualna pozycja karetki w buforze</summary>
    public int CaretOffset
    {
        get => _caretOffset;
        set => _caretOffset = Math.Clamp(value, 0, _buffer.Length);
    }

    /// <summary>Całkowita długość bufora</summary>
    public int Length => _buffer.Length;

    /// <summary>Pełny tekst bufora</summary>
    public string Text => _buffer.ToString();

    /// <summary>Liczba węzłów</summary>
    public int NodeCount => _nodes.Count;

    /// <summary>
    /// Ładuje dokument do wirtualnego bufora (asynchronicznie)
    /// </summary>
    public async Task LoadDocumentAsync(AutomationElement document)
    {
        if (IsLoading || _isRebuilding)
            return;

        IsLoading = true;
        RootElement = document;

        try
        {
            ClearBuffer();

            // Wykryj typ przeglądarki
            DetectBrowser(document);

            Console.WriteLine($"VirtualBuffer: Rozpoczynam ładowanie dokumentu (browser: {ProcessName ?? "unknown"})...");

            await Task.Run(() =>
            {
                BuildBuffer(document);
            });

            Console.WriteLine($"VirtualBuffer: Załadowano {_nodes.Count} węzłów, {_buffer.Length} znaków");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"VirtualBuffer: Błąd ładowania: {ex.Message}");
        }
        finally
        {
            IsLoading = false;
        }
    }

    /// <summary>
    /// Czyści bufor i przygotowuje do ponownego ładowania
    /// </summary>
    private void ClearBuffer()
    {
        _buffer.Clear();
        _nodes.Clear();
        _nodeCache.Clear();
        _nextNodeId = 0;
        _caretOffset = 0;
    }

    /// <summary>
    /// Przebudowuje bufor (np. po zmianie w dokumencie)
    /// Zachowuje bieżącą pozycję karetki jeśli to możliwe
    /// </summary>
    public void Rebuild()
    {
        if (RootElement == null || IsLoading || _isRebuilding)
            return;

        _isRebuilding = true;
        try
        {
            // Zapamiętaj bieżący węzeł na podstawie offset
            var currentNode = GetNodeAtOffset(_caretOffset);
            int savedNodeId = currentNode?.NodeId ?? -1;

            ClearBuffer();
            BuildBuffer(RootElement);

            // Spróbuj przywrócić pozycję karetki
            if (savedNodeId >= 0)
            {
                var restoredNode = _nodes.FirstOrDefault(n => n.NodeId == savedNodeId);
                if (restoredNode != null)
                {
                    _caretOffset = restoredNode.StartOffset;
                }
            }

            Console.WriteLine($"VirtualBuffer: Przebudowano - {_nodes.Count} węzłów");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"VirtualBuffer: Błąd przebudowy: {ex.Message}");
        }
        finally
        {
            _isRebuilding = false;
        }
    }

    /// <summary>
    /// Buduje bufor z drzewa UI Automation
    /// </summary>
    private void BuildBuffer(AutomationElement document)
    {
        TraverseAndBuild(document, 0, null);

        // Zbuduj cache po zakończeniu budowy
        RebuildCache();
    }

    /// <summary>
    /// Przebudowuje cache węzłów dla szybszego dostępu
    /// </summary>
    private void RebuildCache()
    {
        _nodeCache.Clear();
        foreach (var node in _nodes)
        {
            _nodeCache[node.NodeId] = node;
        }
    }

    /// <summary>
    /// Synchronicznie ładuje dokument
    /// </summary>
    public void LoadDocument(AutomationElement document)
    {
        if (IsLoading || _isRebuilding)
            return;

        IsLoading = true;
        RootElement = document;

        try
        {
            ClearBuffer();

            // Wykryj typ przeglądarki
            DetectBrowser(document);

            BuildBuffer(document);

            Console.WriteLine($"VirtualBuffer: Załadowano {_nodes.Count} węzłów, {_buffer.Length} znaków");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"VirtualBuffer: Błąd ładowania: {ex.Message}");
        }
        finally
        {
            IsLoading = false;
        }
    }

    /// <summary>
    /// Wykrywa typ przeglądarki na podstawie procesu
    /// </summary>
    private void DetectBrowser(AutomationElement document)
    {
        try
        {
            var processId = document.Current.ProcessId;
            using var process = System.Diagnostics.Process.GetProcessById(processId);
            ProcessName = process.ProcessName.ToLowerInvariant();
            IsChromiumBrowser = ProcessName is "chrome" or "msedge" or "chromium" or "brave" or "vivaldi" or "opera";

            if (IsChromiumBrowser)
            {
                Console.WriteLine($"VirtualBuffer: Wykryto przeglądarkę Chromium: {ProcessName}");
            }
        }
        catch
        {
            ProcessName = null;
            IsChromiumBrowser = false;
        }
    }

    /// <summary>
    /// Rekurencyjnie buduje bufor z drzewa UI
    /// </summary>
    private void TraverseAndBuild(AutomationElement element, int depth, VirtualBufferNode? parent)
    {
        if (element == null || depth > 100) // Limit głębokości
            return;

        try
        {
            var node = CreateNode(element, parent, depth);

            // Pobierz tekst elementu
            string text = GetNodeText(element);

            if (!string.IsNullOrEmpty(text))
            {
                node.StartOffset = _buffer.Length;
                _buffer.Append(text);

                // Dodaj nową linię po blokowych elementach
                if (IsBlockElement(node.Role))
                {
                    _buffer.Append('\n');
                }

                node.EndOffset = _buffer.Length;
                node.Text = text;
                _nodes.Add(node);
            }

            // Rekurencja dla dzieci
            var walker = TreeWalker.ControlViewWalker;
            var child = walker.GetFirstChild(element);

            while (child != null)
            {
                TraverseAndBuild(child, depth + 1, node);
                child = walker.GetNextSibling(child);
            }
        }
        catch (ElementNotAvailableException)
        {
            // Element zniknął, kontynuuj
        }
        catch (Exception ex)
        {
            Console.WriteLine($"VirtualBuffer: Błąd przy przetwarzaniu węzła: {ex.Message}");
        }
    }

    /// <summary>
    /// Tworzy węzeł dla elementu UI
    /// </summary>
    private VirtualBufferNode CreateNode(AutomationElement element, VirtualBufferNode? parent, int depth)
    {
        var node = new VirtualBufferNode
        {
            Element = element,
            NodeId = _nextNodeId++,
            Parent = parent,
            Depth = depth
        };

        try
        {
            var current = element.Current;
            node.Name = current.Name ?? "";
            node.Role = MapControlTypeToQuickNavType(current.ControlType);
            node.IsFocusable = !current.IsOffscreen && current.IsEnabled;
            node.IsInteractive = IsInteractiveControl(current.ControlType);

            // Pobierz stan elementu
            node.State = GetElementState(element);

            // Pobierz rolę ARIA (dla przeglądarek Chromium)
            if (IsChromiumBrowser)
            {
                var ariaRole = GetAriaRole(element);
                if (!string.IsNullOrEmpty(ariaRole))
                {
                    node.AriaRole = ariaRole;

                    // Mapuj rolę ARIA na QuickNavType
                    var mappedRole = MapAriaRoleToQuickNavType(ariaRole);
                    if (mappedRole != QuickNavType.Text)
                    {
                        node.Role = mappedRole;
                    }
                }

                // Pobierz typ landmark
                var landmarkType = GetLandmarkType(element);
                if (!string.IsNullOrEmpty(landmarkType))
                {
                    node.LandmarkType = landmarkType;
                    if (node.Role == QuickNavType.Text || node.Role == QuickNavType.Group)
                    {
                        node.Role = QuickNavType.Landmark;
                    }
                }

                // Pobierz pozycję w zestawie
                var posInSet = GetPositionInSet(element);
                if (posInSet.HasValue)
                {
                    node.PositionInSet = posInSet.Value.position;
                    node.SizeOfSet = posInSet.Value.size;
                }
            }

            // Pobierz poziom nagłówka
            if (node.Role == QuickNavType.Heading ||
                node.AriaRole?.Equals("heading", StringComparison.OrdinalIgnoreCase) == true)
            {
                node.HeadingLevel = GetHeadingLevel(element);
                if (node.HeadingLevel > 0)
                {
                    node.Role = node.HeadingLevel switch
                    {
                        1 => QuickNavType.Heading1,
                        2 => QuickNavType.Heading2,
                        3 => QuickNavType.Heading3,
                        4 => QuickNavType.Heading4,
                        5 => QuickNavType.Heading5,
                        6 => QuickNavType.Heading6,
                        _ => QuickNavType.Heading
                    };
                }
            }

            // Dodaj do dzieci rodzica
            parent?.Children.Add(node);
        }
        catch { }

        return node;
    }

    /// <summary>
    /// Mapuje rolę ARIA na QuickNavType
    /// </summary>
    private QuickNavType MapAriaRoleToQuickNavType(string ariaRole)
    {
        return ariaRole.ToLowerInvariant() switch
        {
            "heading" => QuickNavType.Heading,
            "link" => QuickNavType.Link,
            "button" => QuickNavType.Button,
            "textbox" or "searchbox" => QuickNavType.EditField,
            "checkbox" => QuickNavType.Checkbox,
            "radio" => QuickNavType.RadioButton,
            "combobox" or "listbox" => QuickNavType.ComboBox,
            "list" => QuickNavType.List,
            "listitem" => QuickNavType.ListItem,
            "table" or "grid" => QuickNavType.Table,
            "row" => QuickNavType.TableRow,
            "cell" or "gridcell" => QuickNavType.TableCell,
            "img" or "image" => QuickNavType.Graphic,
            "main" or "navigation" or "banner" or "contentinfo" or "complementary" or "search" or "form" or "region" => QuickNavType.Landmark,
            "article" => QuickNavType.Article,
            "blockquote" => QuickNavType.BlockQuote,
            "separator" => QuickNavType.Separator,
            "menu" => QuickNavType.Menu,
            "menuitem" => QuickNavType.MenuItem,
            "tab" => QuickNavType.Tab,
            "tabpanel" => QuickNavType.TabPanel,
            "tree" => QuickNavType.Tree,
            "treeitem" => QuickNavType.TreeItem,
            "alert" or "alertdialog" => QuickNavType.Alert,
            "dialog" => QuickNavType.Dialog,
            "progressbar" => QuickNavType.ProgressBar,
            "slider" => QuickNavType.Slider,
            _ => QuickNavType.Text
        };
    }

    /// <summary>
    /// Pobiera tekst z elementu
    /// </summary>
    private string GetNodeText(AutomationElement element)
    {
        try
        {
            // Próbuj TextPattern
            if (element.TryGetCurrentPattern(TextPattern.Pattern, out var textPattern))
            {
                return ((TextPattern)textPattern).DocumentRange.GetText(-1);
            }

            // Próbuj ValuePattern
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var valuePattern))
            {
                return ((ValuePattern)valuePattern).Current.Value;
            }

            // Użyj nazwy
            var name = element.Current.Name;
            if (!string.IsNullOrEmpty(name))
                return name;
        }
        catch { }

        return "";
    }

    /// <summary>
    /// Mapuje typ kontrolki na QuickNavType
    /// </summary>
    private QuickNavType MapControlTypeToQuickNavType(ControlType controlType)
    {
        if (controlType == ControlType.Hyperlink)
            return QuickNavType.Link;
        if (controlType == ControlType.Button)
            return QuickNavType.Button;
        if (controlType == ControlType.Edit)
            return QuickNavType.EditField;
        if (controlType == ControlType.CheckBox)
            return QuickNavType.Checkbox;
        if (controlType == ControlType.RadioButton)
            return QuickNavType.RadioButton;
        if (controlType == ControlType.ComboBox)
            return QuickNavType.ComboBox;
        if (controlType == ControlType.List)
            return QuickNavType.List;
        if (controlType == ControlType.ListItem)
            return QuickNavType.ListItem;
        if (controlType == ControlType.Table)
            return QuickNavType.Table;
        if (controlType == ControlType.DataItem)
            return QuickNavType.TableCell;
        if (controlType == ControlType.Image)
            return QuickNavType.Graphic;
        if (controlType == ControlType.Group)
            return QuickNavType.Landmark;
        if (controlType == ControlType.Document)
            return QuickNavType.Document;
        if (controlType == ControlType.Text)
            return QuickNavType.Text;
        if (controlType == ControlType.Header || controlType == ControlType.HeaderItem)
            return QuickNavType.Heading;

        return QuickNavType.Text;
    }

    /// <summary>
    /// Sprawdza czy element jest interaktywny
    /// </summary>
    private bool IsInteractiveControl(ControlType controlType)
    {
        return controlType == ControlType.Button ||
               controlType == ControlType.Edit ||
               controlType == ControlType.Hyperlink ||
               controlType == ControlType.CheckBox ||
               controlType == ControlType.RadioButton ||
               controlType == ControlType.ComboBox ||
               controlType == ControlType.MenuItem ||
               controlType == ControlType.TabItem;
    }

    /// <summary>
    /// Sprawdza czy to element blokowy (wymaga nowej linii)
    /// </summary>
    private bool IsBlockElement(QuickNavType type)
    {
        return type is QuickNavType.Heading or QuickNavType.Heading1 or
            QuickNavType.Heading2 or QuickNavType.Heading3 or
            QuickNavType.Heading4 or QuickNavType.Heading5 or
            QuickNavType.Heading6 or QuickNavType.Paragraph or
            QuickNavType.List or QuickNavType.ListItem or
            QuickNavType.Table or QuickNavType.BlockQuote;
    }

    /// <summary>
    /// Pobiera stan elementu
    /// </summary>
    private ElementState GetElementState(AutomationElement element)
    {
        ElementState state = ElementState.None;

        try
        {
            if (element.TryGetCurrentPattern(TogglePattern.Pattern, out var togglePattern))
            {
                if (((TogglePattern)togglePattern).Current.ToggleState == ToggleState.On)
                    state |= ElementState.Checked;
            }

            if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selectionPattern))
            {
                if (((SelectionItemPattern)selectionPattern).Current.IsSelected)
                    state |= ElementState.Selected;
            }

            if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var expandPattern))
            {
                var expandState = ((ExpandCollapsePattern)expandPattern).Current.ExpandCollapseState;
                if (expandState == ExpandCollapseState.Expanded)
                    state |= ElementState.Expanded;
                else if (expandState == ExpandCollapseState.Collapsed)
                    state |= ElementState.Collapsed;
            }

            if (!element.Current.IsEnabled)
                state |= ElementState.Disabled;
        }
        catch { }

        return state;
    }

    /// <summary>
    /// Pobiera poziom nagłówka
    /// </summary>
    private int GetHeadingLevel(AutomationElement element)
    {
        try
        {
            // 1. Próbuj pobrać z UIA Level Property (najlepsze dla Chromium)
            if (AriaUiaProperties.LevelProperty != null)
            {
                var levelObj = element.GetCurrentPropertyValue(AriaUiaProperties.LevelProperty);
                if (levelObj is int level && level >= 1 && level <= 6)
                {
                    return level;
                }
            }

            // 2. Próbuj pobrać z AriaRole
            if (AriaUiaProperties.AriaRoleProperty != null)
            {
                var ariaRole = element.GetCurrentPropertyValue(AriaUiaProperties.AriaRoleProperty) as string;
                if (!string.IsNullOrEmpty(ariaRole))
                {
                    // ARIA role może być "heading" z aria-level
                    if (ariaRole.Equals("heading", StringComparison.OrdinalIgnoreCase))
                    {
                        // Sprawdź aria-level w AriaProperties
                        if (AriaUiaProperties.AriaPropertiesProperty != null)
                        {
                            var ariaProps = element.GetCurrentPropertyValue(AriaUiaProperties.AriaPropertiesProperty) as string;
                            if (!string.IsNullOrEmpty(ariaProps))
                            {
                                // Format: "level=2;..."
                                var match = System.Text.RegularExpressions.Regex.Match(ariaProps, @"level\s*=\s*(\d+)");
                                if (match.Success && int.TryParse(match.Groups[1].Value, out int ariaLevel))
                                {
                                    return Math.Clamp(ariaLevel, 1, 6);
                                }
                            }
                        }
                    }

                    // Sprawdź role h1-h6
                    for (int i = 1; i <= 6; i++)
                    {
                        if (ariaRole.Equals($"h{i}", StringComparison.OrdinalIgnoreCase) ||
                            ariaRole.Equals($"heading{i}", StringComparison.OrdinalIgnoreCase))
                        {
                            return i;
                        }
                    }
                }
            }

            // 3. Sprawdź LocalizedControlType
            var localizedType = element.Current.LocalizedControlType?.ToLowerInvariant();
            if (!string.IsNullOrEmpty(localizedType))
            {
                // Szukaj wzorca "nagłówek X" lub "heading X"
                var match = System.Text.RegularExpressions.Regex.Match(localizedType, @"(?:nagłówek|heading)\s*(\d+)");
                if (match.Success && int.TryParse(match.Groups[1].Value, out int typeLevel))
                {
                    return Math.Clamp(typeLevel, 1, 6);
                }
            }

            // 4. Fallback - sprawdź nazwę elementu
            var name = element.Current.Name;
            if (!string.IsNullOrEmpty(name))
            {
                if (name.StartsWith("heading level", StringComparison.OrdinalIgnoreCase))
                {
                    if (int.TryParse(name.AsSpan(14), out int level))
                        return Math.Clamp(level, 1, 6);
                }

                // Szukaj wzorca "poziom X" lub "level X" w nazwie
                var match = System.Text.RegularExpressions.Regex.Match(name, @"(?:level|poziom)\s*(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (match.Success && int.TryParse(match.Groups[1].Value, out int nameLevel))
                {
                    return Math.Clamp(nameLevel, 1, 6);
                }
            }
        }
        catch { }

        return 0;
    }

    /// <summary>
    /// Pobiera rolę ARIA elementu
    /// </summary>
    private string? GetAriaRole(AutomationElement element)
    {
        try
        {
            if (AriaUiaProperties.AriaRoleProperty != null)
            {
                return element.GetCurrentPropertyValue(AriaUiaProperties.AriaRoleProperty) as string;
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Pobiera typ landmark (region ARIA)
    /// </summary>
    private string? GetLandmarkType(AutomationElement element)
    {
        try
        {
            var ariaRole = GetAriaRole(element);
            if (!string.IsNullOrEmpty(ariaRole))
            {
                var role = ariaRole.ToLowerInvariant();
                return role switch
                {
                    "main" => "główny",
                    "navigation" or "nav" => "nawigacja",
                    "banner" => "baner",
                    "contentinfo" => "informacje o treści",
                    "complementary" or "aside" => "uzupełniający",
                    "search" => "wyszukiwarka",
                    "form" => "formularz",
                    "region" => "region",
                    "article" => "artykuł",
                    "application" => "aplikacja",
                    _ => null
                };
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Pobiera pozycję elementu w zestawie (np. "2 z 5")
    /// </summary>
    private (int position, int size)? GetPositionInSet(AutomationElement element)
    {
        try
        {
            int position = 0;
            int size = 0;

            if (AriaUiaProperties.PositionInSetProperty != null)
            {
                var posObj = element.GetCurrentPropertyValue(AriaUiaProperties.PositionInSetProperty);
                if (posObj is int pos)
                    position = pos;
            }

            if (AriaUiaProperties.SizeOfSetProperty != null)
            {
                var sizeObj = element.GetCurrentPropertyValue(AriaUiaProperties.SizeOfSetProperty);
                if (sizeObj is int sz)
                    size = sz;
            }

            if (position > 0 && size > 0)
                return (position, size);
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Znajduje następny węzeł danego typu
    /// </summary>
    public VirtualBufferNode? FindNextByType(QuickNavType type, int fromOffset)
    {
        return _nodes
            .Where(n => n.StartOffset > fromOffset && MatchesType(n.Role, type))
            .OrderBy(n => n.StartOffset)
            .FirstOrDefault();
    }

    /// <summary>
    /// Znajduje poprzedni węzeł danego typu
    /// </summary>
    public VirtualBufferNode? FindPreviousByType(QuickNavType type, int fromOffset)
    {
        return _nodes
            .Where(n => n.EndOffset < fromOffset && MatchesType(n.Role, type))
            .OrderByDescending(n => n.StartOffset)
            .FirstOrDefault();
    }

    /// <summary>
    /// Sprawdza czy typy pasują (z uwzględnieniem hierarchii)
    /// </summary>
    private bool MatchesType(QuickNavType nodeType, QuickNavType searchType)
    {
        if (nodeType == searchType)
            return true;

        // Nagłówki
        if (searchType == QuickNavType.Heading && QuickNavKeys.IsHeading(nodeType))
            return true;

        // Linki
        if (searchType == QuickNavType.Link && QuickNavKeys.IsLink(nodeType))
            return true;

        // Pola formularzy
        if (searchType == QuickNavType.FormField && QuickNavKeys.IsFormField(nodeType))
            return true;

        return false;
    }

    /// <summary>
    /// Pobiera węzeł na danym offsecie
    /// Używa binarnego wyszukiwania dla wydajności
    /// </summary>
    public VirtualBufferNode? GetNodeAtOffset(int offset)
    {
        if (_nodes.Count == 0)
            return null;

        // Binary search dla wydajności - węzły są posortowane po StartOffset
        int left = 0;
        int right = _nodes.Count - 1;

        while (left <= right)
        {
            int mid = left + (right - left) / 2;
            var node = _nodes[mid];

            if (node.ContainsOffset(offset))
                return node;

            if (offset < node.StartOffset)
                right = mid - 1;
            else
                left = mid + 1;
        }

        return null;
    }

    /// <summary>
    /// Pobiera węzeł po ID z cache
    /// </summary>
    public VirtualBufferNode? GetNodeById(int nodeId)
    {
        return _nodeCache.TryGetValue(nodeId, out var node) ? node : null;
    }

    /// <summary>
    /// Pobiera tekst w zakresie
    /// </summary>
    public string GetTextRange(int start, int end)
    {
        start = Math.Clamp(start, 0, _buffer.Length);
        end = Math.Clamp(end, start, _buffer.Length);

        if (start >= end)
            return "";

        return _buffer.ToString(start, end - start);
    }

    /// <summary>
    /// Oblicza offset początku linii dla danego offsetu
    /// </summary>
    public int GetLineStartOffset(int offset)
    {
        offset = Math.Clamp(offset, 0, _buffer.Length);

        int start = offset;
        while (start > 0 && _buffer[start - 1] != '\n')
            start--;

        return start;
    }

    /// <summary>
    /// Oblicza offset końca linii dla danego offsetu
    /// </summary>
    public int GetLineEndOffset(int offset)
    {
        offset = Math.Clamp(offset, 0, _buffer.Length);

        int end = offset;
        while (end < _buffer.Length && _buffer[end] != '\n')
            end++;

        return end;
    }

    /// <summary>
    /// Pobiera zakres linii dla danego offsetu
    /// </summary>
    public (int start, int end) GetLineRange(int offset)
    {
        int start = GetLineStartOffset(offset);
        int end = GetLineEndOffset(offset);
        return (start, end);
    }

    /// <summary>
    /// Znajduje offset początku słowa zawierającego dany offset
    /// </summary>
    public int GetWordStartOffset(int offset)
    {
        offset = Math.Clamp(offset, 0, _buffer.Length);

        int start = offset;
        while (start > 0 && !char.IsWhiteSpace(_buffer[start - 1]))
            start--;

        return start;
    }

    /// <summary>
    /// Znajduje offset końca słowa zawierającego dany offset
    /// </summary>
    public int GetWordEndOffset(int offset)
    {
        offset = Math.Clamp(offset, 0, _buffer.Length);

        int end = offset;
        while (end < _buffer.Length && !char.IsWhiteSpace(_buffer[end]))
            end++;

        return end;
    }

    /// <summary>
    /// Pobiera zakres słowa dla danego offsetu
    /// </summary>
    public (int start, int end) GetWordRange(int offset)
    {
        int start = GetWordStartOffset(offset);
        int end = GetWordEndOffset(offset);
        return (start, end);
    }

    /// <summary>
    /// Pobiera bieżącą linię na podstawie pozycji karetki
    /// </summary>
    public (int start, int end, string text) GetCurrentLine()
    {
        if (_buffer.Length == 0)
            return (0, 0, "");

        var (start, end) = GetLineRange(_caretOffset);
        string text = GetTextRange(start, end);
        return (start, end, text);
    }

    /// <summary>
    /// Przesuwa do następnego znaku i zwraca jego opis fonetyczny
    /// </summary>
    public string MoveToNextCharacter()
    {
        if (_caretOffset < _buffer.Length - 1)
        {
            _caretOffset++;
        }
        else if (_caretOffset >= _buffer.Length)
        {
            return "Koniec dokumentu";
        }

        return GetCharacterAtOffset(_caretOffset);
    }

    /// <summary>
    /// Przesuwa do poprzedniego znaku i zwraca jego opis fonetyczny
    /// </summary>
    public string MoveToPreviousCharacter()
    {
        if (_caretOffset > 0)
        {
            _caretOffset--;
        }
        else
        {
            return "Początek dokumentu";
        }

        return GetCharacterAtOffset(_caretOffset);
    }

    /// <summary>
    /// Pobiera opis fonetyczny znaku na danej pozycji
    /// </summary>
    public string GetCharacterAtOffset(int offset)
    {
        if (offset < 0 || offset >= _buffer.Length)
            return "";

        char ch = _buffer[offset];
        return GetPhoneticDescription(ch);
    }

    /// <summary>
    /// Pobiera bieżący znak (bez przesuwania)
    /// </summary>
    public string GetCurrentCharacter()
    {
        return GetCharacterAtOffset(_caretOffset);
    }

    /// <summary>
    /// Zwraca fonetyczny opis znaku (po polsku)
    /// </summary>
    private static string GetPhoneticDescription(char ch)
    {
        return char.ToUpperInvariant(ch) switch
        {
            ' ' => "spacja",
            '\n' => "nowa linia",
            '\r' => "powrót karetki",
            '\t' => "tabulator",
            '.' => "kropka",
            ',' => "przecinek",
            ';' => "średnik",
            ':' => "dwukropek",
            '!' => "wykrzyknik",
            '?' => "znak zapytania",
            '-' => "myślnik",
            '_' => "podkreślenie",
            '(' => "nawias otwierający",
            ')' => "nawias zamykający",
            '[' => "nawias kwadratowy otwierający",
            ']' => "nawias kwadratowy zamykający",
            '{' => "nawias klamrowy otwierający",
            '}' => "nawias klamrowy zamykający",
            '<' => "mniejszy niż",
            '>' => "większy niż",
            '/' => "ukośnik",
            '\\' => "ukośnik odwrotny",
            '@' => "małpa",
            '#' => "hash",
            '$' => "dolar",
            '%' => "procent",
            '^' => "daszek",
            '&' => "ampersand",
            '*' => "gwiazdka",
            '+' => "plus",
            '=' => "równa się",
            '"' => "cudzysłów",
            '\'' => "apostrof",
            '`' => "grawis",
            '~' => "tylda",
            '|' => "kreska pionowa",
            // Wielkie litery
            >= 'A' and <= 'Z' => $"duże {ch}",
            // Małe litery i inne
            _ => ch.ToString()
        };
    }

    /// <summary>
    /// Przesuwa do następnego słowa i zwraca je
    /// </summary>
    public string MoveToNextWord()
    {
        if (_caretOffset >= _buffer.Length)
            return "Koniec dokumentu";

        // Pomiń bieżące słowo (znaki nie-białe)
        while (_caretOffset < _buffer.Length && !char.IsWhiteSpace(_buffer[_caretOffset]))
            _caretOffset++;

        // Pomiń białe znaki
        while (_caretOffset < _buffer.Length && char.IsWhiteSpace(_buffer[_caretOffset]))
            _caretOffset++;

        if (_caretOffset >= _buffer.Length)
            return "Koniec dokumentu";

        return GetWordAtOffset(_caretOffset);
    }

    /// <summary>
    /// Przesuwa do poprzedniego słowa i zwraca je
    /// </summary>
    public string MoveToPreviousWord()
    {
        if (_caretOffset <= 0)
            return "Początek dokumentu";

        // Cofnij o jeden znak
        _caretOffset--;

        // Pomiń białe znaki wstecz
        while (_caretOffset > 0 && char.IsWhiteSpace(_buffer[_caretOffset]))
            _caretOffset--;

        // Znajdź początek słowa
        while (_caretOffset > 0 && !char.IsWhiteSpace(_buffer[_caretOffset - 1]))
            _caretOffset--;

        return GetWordAtOffset(_caretOffset);
    }

    /// <summary>
    /// Pobiera słowo na danej pozycji
    /// </summary>
    public string GetWordAtOffset(int offset)
    {
        if (offset < 0 || offset >= _buffer.Length)
            return "";

        var (start, end) = GetWordRange(offset);
        return GetTextRange(start, end);
    }

    /// <summary>
    /// Pobiera bieżące słowo (bez przesuwania)
    /// </summary>
    public string GetCurrentWord()
    {
        if (_buffer.Length == 0 || _caretOffset >= _buffer.Length)
            return "";

        var (start, _) = GetWordRange(_caretOffset);
        return GetWordAtOffset(start);
    }

    /// <summary>
    /// Przesuwa do następnej linii i zwraca jej tekst
    /// </summary>
    public string MoveToNextLine()
    {
        var (_, end, _) = GetCurrentLine();
        if (end < _buffer.Length)
        {
            _caretOffset = end + 1;
            var (_, _, text) = GetCurrentLine();
            return string.IsNullOrWhiteSpace(text) ? "pusta linia" : text;
        }
        return "Koniec dokumentu";
    }

    /// <summary>
    /// Przesuwa do poprzedniej linii i zwraca jej tekst
    /// </summary>
    public string MoveToPreviousLine()
    {
        var (start, _, _) = GetCurrentLine();
        if (start > 0)
        {
            _caretOffset = start - 1;
            var (newStart, _, text) = GetCurrentLine();
            _caretOffset = newStart;
            return string.IsNullOrWhiteSpace(text) ? "pusta linia" : text;
        }
        return "Początek dokumentu";
    }

    /// <summary>
    /// Przesuwa karetkę do węzła
    /// </summary>
    public void MoveToNode(VirtualBufferNode node)
    {
        _caretOffset = node.StartOffset;
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _nodes.Clear();
        _buffer.Clear();
        RootElement = null;
        _disposed = true;
    }
}
