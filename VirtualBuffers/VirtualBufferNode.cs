using System.Windows.Automation;
using ScreenReader.BrowseMode;

namespace ScreenReader.VirtualBuffers;

/// <summary>
/// Reprezentuje węzeł w wirtualnym buforze
/// Port z NVDA virtualBuffers - struktura węzła
/// </summary>
public class VirtualBufferNode
{
    /// <summary>Element UI Automation</summary>
    public AutomationElement? Element { get; set; }

    /// <summary>Offset początkowy w buforze tekstowym</summary>
    public int StartOffset { get; set; }

    /// <summary>Offset końcowy w buforze tekstowym</summary>
    public int EndOffset { get; set; }

    /// <summary>Rola/typ elementu</summary>
    public QuickNavType Role { get; set; }

    /// <summary>Tekst węzła</summary>
    public string Text { get; set; } = "";

    /// <summary>Nazwa elementu</summary>
    public string Name { get; set; } = "";

    /// <summary>Wartość elementu (dla pól edycyjnych, itp.)</summary>
    public string Value { get; set; } = "";

    /// <summary>Opis elementu</summary>
    public string Description { get; set; } = "";

    /// <summary>Poziom nagłówka (1-6, 0 jeśli nie nagłówek)</summary>
    public int HeadingLevel { get; set; }

    /// <summary>Czy element jest interaktywny</summary>
    public bool IsInteractive { get; set; }

    /// <summary>Czy element jest fokusowy</summary>
    public bool IsFocusable { get; set; }

    /// <summary>Stan elementu (zaznaczony, rozwinięty, itp.)</summary>
    public ElementState State { get; set; }

    /// <summary>Atrybuty ARIA</summary>
    public Dictionary<string, string> AriaAttributes { get; set; } = new();

    /// <summary>Rola ARIA elementu (np. "heading", "button", "link")</summary>
    public string? AriaRole { get; set; }

    /// <summary>Typ landmarku ARIA (np. "główny", "nawigacja")</summary>
    public string? LandmarkType { get; set; }

    /// <summary>Pozycja w zestawie (np. 2 z 5)</summary>
    public int PositionInSet { get; set; }

    /// <summary>Rozmiar zestawu</summary>
    public int SizeOfSet { get; set; }

    /// <summary>ID dokumentu (dla identyfikacji)</summary>
    public int DocHandle { get; set; }

    /// <summary>Unikalny identyfikator węzła</summary>
    public int NodeId { get; set; }

    /// <summary>Węzeł nadrzędny</summary>
    public VirtualBufferNode? Parent { get; set; }

    /// <summary>Węzły potomne</summary>
    public List<VirtualBufferNode> Children { get; set; } = new();

    /// <summary>Głębokość w drzewie</summary>
    public int Depth { get; set; }

    /// <summary>Długość tekstu węzła</summary>
    public int Length => EndOffset - StartOffset;

    /// <summary>Czy offset znajduje się w tym węźle</summary>
    public bool ContainsOffset(int offset)
    {
        return offset >= StartOffset && offset < EndOffset;
    }

    /// <summary>Pobiera ogłoszenie dla węzła</summary>
    public string GetAnnouncement()
    {
        var parts = new List<string>();

        // Nazwa
        if (!string.IsNullOrEmpty(Name))
            parts.Add(Name);
        else if (!string.IsNullOrEmpty(Text))
            parts.Add(Text);

        // Rola
        string roleText = GetRoleText();
        if (!string.IsNullOrEmpty(roleText))
            parts.Add(roleText);

        // Typ landmarku (jeśli jest)
        if (!string.IsNullOrEmpty(LandmarkType))
            parts.Add(LandmarkType);

        // Poziom nagłówka
        if (HeadingLevel > 0)
            parts.Add($"poziom {HeadingLevel}");

        // Pozycja w zestawie
        if (PositionInSet > 0 && SizeOfSet > 0)
            parts.Add($"{PositionInSet} z {SizeOfSet}");

        // Stan
        string stateText = GetStateText();
        if (!string.IsNullOrEmpty(stateText))
            parts.Add(stateText);

        // Wartość
        if (!string.IsNullOrEmpty(Value))
            parts.Add(Value);

        return string.Join(", ", parts);
    }

    /// <summary>Pobiera tekst roli po polsku</summary>
    private string GetRoleText()
    {
        return Role switch
        {
            QuickNavType.Heading => "nagłówek",
            QuickNavType.Link => "link",
            QuickNavType.Button => "przycisk",
            QuickNavType.EditField => "pole edycji",
            QuickNavType.Checkbox => "pole wyboru",
            QuickNavType.RadioButton => "przycisk opcji",
            QuickNavType.ComboBox => "pole kombi",
            QuickNavType.List => "lista",
            QuickNavType.ListItem => "element listy",
            QuickNavType.Table => "tabela",
            QuickNavType.TableCell => "komórka",
            QuickNavType.Graphic => "grafika",
            QuickNavType.Landmark => "obszar",
            QuickNavType.FormField => "pole formularza",
            QuickNavType.Frame => "ramka",
            QuickNavType.BlockQuote => "cytat",
            _ => ""
        };
    }

    /// <summary>Pobiera tekst stanu po polsku</summary>
    private string GetStateText()
    {
        var states = new List<string>();

        if (State.HasFlag(ElementState.Checked))
            states.Add("zaznaczone");
        if (State.HasFlag(ElementState.Pressed))
            states.Add("wciśnięty");
        if (State.HasFlag(ElementState.Expanded))
            states.Add("rozwinięty");
        if (State.HasFlag(ElementState.Collapsed))
            states.Add("zwinięty");
        if (State.HasFlag(ElementState.Selected))
            states.Add("wybrane");
        if (State.HasFlag(ElementState.Visited))
            states.Add("odwiedzony");
        if (State.HasFlag(ElementState.Required))
            states.Add("wymagane");
        if (State.HasFlag(ElementState.Invalid))
            states.Add("nieprawidłowe");
        if (State.HasFlag(ElementState.ReadOnly))
            states.Add("tylko do odczytu");
        if (State.HasFlag(ElementState.Disabled))
            states.Add("niedostępny");

        return string.Join(", ", states);
    }

    public override string ToString()
    {
        return $"[{StartOffset}-{EndOffset}] {Role}: {Name ?? Text}";
    }
}

/// <summary>
/// Stan elementu (bitflags)
/// </summary>
[Flags]
public enum ElementState
{
    None = 0,
    Checked = 1,
    Pressed = 2,
    Expanded = 4,
    Collapsed = 8,
    Selected = 16,
    Visited = 32,
    Required = 64,
    Invalid = 128,
    ReadOnly = 256,
    Disabled = 512,
    HasPopup = 1024,
    Busy = 2048
}
