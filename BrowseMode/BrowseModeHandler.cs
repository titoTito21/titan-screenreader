using System.Windows.Automation;
using ScreenReader.VirtualBuffers;

namespace ScreenReader.BrowseMode;

/// <summary>
/// Obsługuje tryb przeglądania (browse mode) dla dokumentów webowych
/// Port z NVDA browseMode.py
/// </summary>
public class BrowseModeHandler : IDisposable
{
    private VirtualBuffer? _buffer;
    private bool _passThrough;
    private bool _virtualCursorEnabled = true;
    private bool _disposed;

    /// <summary>Event wywoływany gdy należy ogłosić element</summary>
    public event Action<string>? Announce;

    /// <summary>Event wywoływany przy zmianie trybu</summary>
    public event Action<bool>? ModeChanged;

    /// <summary>Aktywny wirtualny bufor</summary>
    public VirtualBuffer? Buffer => _buffer;

    /// <summary>Czy tryb pass-through (focus mode) jest aktywny</summary>
    public bool PassThrough
    {
        get => _passThrough;
        private set
        {
            if (_passThrough != value)
            {
                _passThrough = value;
                if (_buffer != null)
                    _buffer.PassThrough = value;
                ModeChanged?.Invoke(value);
            }
        }
    }

    /// <summary>Czy browse mode jest aktywny</summary>
    public bool IsActive => _buffer != null && !_buffer.IsLoading;

    /// <summary>Czy wirtualny kursor (TCE) jest włączony</summary>
    public bool IsVirtualCursorEnabled
    {
        get => _virtualCursorEnabled;
        private set => _virtualCursorEnabled = value;
    }

    /// <summary>
    /// Przełącza wirtualny kursor (TCE) i zwraca komunikat
    /// </summary>
    /// <returns>"Kursor TCE tak" lub "Kursor TCE nie"</returns>
    public string ToggleVirtualCursor()
    {
        IsVirtualCursorEnabled = !IsVirtualCursorEnabled;
        return IsVirtualCursorEnabled ? "Kursor TCE tak" : "Kursor TCE nie";
    }

    /// <summary>
    /// Aktywuje browse mode dla dokumentu
    /// </summary>
    public async Task ActivateAsync(AutomationElement document)
    {
        _buffer?.Dispose();
        _buffer = new VirtualBuffer();
        _passThrough = false;

        Announce?.Invoke("Ładowanie dokumentu...");

        await _buffer.LoadDocumentAsync(document);

        if (_buffer.Length > 0)
        {
            var firstNode = _buffer.GetNodeAtOffset(0);
            if (firstNode != null)
            {
                AnnounceNode(firstNode);
            }
        }

        Announce?.Invoke($"Dokument załadowany, {_buffer.NodeCount} elementów");
    }

    /// <summary>
    /// Aktywuje browse mode synchronicznie
    /// </summary>
    public void Activate(AutomationElement document)
    {
        _buffer?.Dispose();
        _buffer = new VirtualBuffer();
        _passThrough = false;

        _buffer.LoadDocument(document);

        if (_buffer.Length > 0)
        {
            var firstNode = _buffer.GetNodeAtOffset(0);
            if (firstNode != null)
            {
                AnnounceNode(firstNode);
            }
        }
    }

    /// <summary>
    /// Dezaktywuje browse mode
    /// </summary>
    public void Deactivate()
    {
        _buffer?.Dispose();
        _buffer = null;
        _passThrough = false;
    }

    /// <summary>
    /// Przełącza między browse mode a focus mode
    /// </summary>
    public void TogglePassThrough()
    {
        PassThrough = !PassThrough;

        string modeName = PassThrough ? "Tryb formularza" : "Tryb przeglądania";
        Announce?.Invoke(modeName);
    }

    /// <summary>
    /// Obsługuje szybką nawigację jednoliterową
    /// </summary>
    /// <param name="key">Klawisz nawigacji</param>
    /// <param name="shift">Czy Shift jest wciśnięty (nawigacja wstecz)</param>
    /// <returns>True jeśli nawigacja została obsłużona</returns>
    public bool HandleQuickNav(char key, bool shift)
    {
        if (_buffer == null || PassThrough || !_virtualCursorEnabled)
            return false;

        var type = QuickNavKeys.GetTypeForKey(key);
        if (type == QuickNavType.None)
            return false;

        VirtualBufferNode? node;

        if (shift)
        {
            node = _buffer.FindPreviousByType(type, _buffer.CaretOffset);
            if (node == null)
            {
                Announce?.Invoke($"Brak poprzedniego elementu typu {QuickNavKeys.GetTypeName(type)}");
                return true;
            }
        }
        else
        {
            node = _buffer.FindNextByType(type, _buffer.CaretOffset);
            if (node == null)
            {
                Announce?.Invoke($"Brak następnego elementu typu {QuickNavKeys.GetTypeName(type)}");
                return true;
            }
        }

        MoveTo(node);
        AnnounceNode(node);
        return true;
    }

    /// <summary>
    /// Przesuwa do węzła
    /// </summary>
    private void MoveTo(VirtualBufferNode node)
    {
        if (_buffer == null)
            return;

        _buffer.MoveToNode(node);
    }

    /// <summary>
    /// Ogłasza węzeł
    /// </summary>
    private void AnnounceNode(VirtualBufferNode node)
    {
        string announcement = node.GetAnnouncement();
        Announce?.Invoke(announcement);
    }

    /// <summary>
    /// Przesuwa do następnej linii
    /// </summary>
    public void MoveToNextLine()
    {
        if (_buffer == null || PassThrough)
            return;

        var (_, end, _) = _buffer.GetCurrentLine();
        if (end < _buffer.Length)
        {
            _buffer.CaretOffset = end + 1;
            ReadCurrentLine();
        }
        else
        {
            Announce?.Invoke("Koniec dokumentu");
        }
    }

    /// <summary>
    /// Przesuwa do poprzedniej linii
    /// </summary>
    public void MoveToPreviousLine()
    {
        if (_buffer == null || PassThrough)
            return;

        var (start, _, _) = _buffer.GetCurrentLine();
        if (start > 0)
        {
            _buffer.CaretOffset = start - 1;
            var (newStart, _, _) = _buffer.GetCurrentLine();
            _buffer.CaretOffset = newStart;
            ReadCurrentLine();
        }
        else
        {
            Announce?.Invoke("Początek dokumentu");
        }
    }

    /// <summary>
    /// Czyta bieżącą linię
    /// </summary>
    public void ReadCurrentLine()
    {
        if (_buffer == null)
            return;

        var (_, _, text) = _buffer.GetCurrentLine();
        if (!string.IsNullOrWhiteSpace(text))
        {
            Announce?.Invoke(text);
        }
        else
        {
            Announce?.Invoke("pusta linia");
        }
    }

    /// <summary>
    /// Czyta bieżący element
    /// </summary>
    public void ReadCurrentElement()
    {
        if (_buffer == null)
            return;

        var node = _buffer.GetNodeAtOffset(_buffer.CaretOffset);
        if (node != null)
        {
            AnnounceNode(node);
        }
    }

    /// <summary>
    /// Aktywuje bieżący element (Enter/Space)
    /// </summary>
    public bool ActivateCurrentElement()
    {
        if (_buffer == null)
            return false;

        var node = _buffer.GetNodeAtOffset(_buffer.CaretOffset);
        if (node?.Element == null)
            return false;

        // Dla interaktywnych elementów, przełącz na focus mode
        if (node.IsInteractive)
        {
            try
            {
                // Spróbuj kliknąć lub ustawić fokus
                if (node.Element.TryGetCurrentPattern(InvokePattern.Pattern, out var invokePattern))
                {
                    ((InvokePattern)invokePattern).Invoke();
                    return true;
                }

                // Dla pól edycyjnych przełącz na focus mode
                if (node.Role == QuickNavType.EditField)
                {
                    PassThrough = true;
                    node.Element.SetFocus();
                    return true;
                }

                // Spróbuj ustawić fokus
                node.Element.SetFocus();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BrowseModeHandler: Błąd aktywacji: {ex.Message}");
            }
        }

        return false;
    }

    /// <summary>
    /// Przesuwa do początku dokumentu
    /// </summary>
    public void MoveToStart()
    {
        if (_buffer == null || PassThrough)
            return;

        _buffer.CaretOffset = 0;
        ReadCurrentLine();
        Announce?.Invoke("Początek dokumentu");
    }

    /// <summary>
    /// Przesuwa do końca dokumentu
    /// </summary>
    public void MoveToEnd()
    {
        if (_buffer == null || PassThrough)
            return;

        _buffer.CaretOffset = _buffer.Length;
        ReadCurrentLine();
        Announce?.Invoke("Koniec dokumentu");
    }

    /// <summary>
    /// Czyta cały dokument
    /// </summary>
    public void SayAll()
    {
        if (_buffer == null)
            return;

        Announce?.Invoke(_buffer.Text);
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _buffer?.Dispose();
        _disposed = true;
    }
}
