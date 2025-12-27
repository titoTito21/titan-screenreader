using System.Windows.Forms;

namespace ScreenReader;

/// <summary>
/// Menu kontekstowe czytnika ekranu
/// Wyświetla się na pozycji kursora i zostaje na pierwszym planie
/// </summary>
public class ScreenReaderContextMenu : ContextMenuStrip
{
    private readonly Action? _onSettings;
    private readonly Action? _onHelp;
    private readonly Action? _onExit;

    public ScreenReaderContextMenu(Action? onSettings, Action? onHelp, Action? onExit)
    {
        _onSettings = onSettings;
        _onHelp = onHelp;
        _onExit = onExit;

        InitializeItems();

        // Upewnij się, że menu jest na pierwszym planie
        TopLevel = true;
    }

    private void InitializeItems()
    {
        // Ustawienia
        var settingsItem = new ToolStripMenuItem("&Ustawienia...")
        {
            ShortcutKeyDisplayString = "Insert+N, U"
        };
        settingsItem.Click += (s, e) => _onSettings?.Invoke();

        // Pomoc
        var helpItem = new ToolStripMenuItem("&Pomoc")
        {
            ShortcutKeyDisplayString = "Insert+F1"
        };
        helpItem.Click += (s, e) => _onHelp?.Invoke();

        // Separator
        var separator = new ToolStripSeparator();

        // Wyjście
        var exitItem = new ToolStripMenuItem("&Zamknij czytnik ekranu")
        {
            ShortcutKeyDisplayString = "Alt+F4"
        };
        exitItem.Click += (s, e) => _onExit?.Invoke();

        Items.AddRange(new ToolStripItem[]
        {
            settingsItem,
            helpItem,
            separator,
            exitItem
        });
    }

    /// <summary>
    /// Pokazuje menu na pozycji kursora myszy
    /// </summary>
    public void ShowAtCursor()
    {
        // Pokaż menu na aktualnej pozycji kursora
        Show(Cursor.Position);

        // Ustaw fokus na menu
        Focus();
    }

    /// <summary>
    /// Pokazuje menu na środku ekranu (dla użytkowników bez myszy)
    /// </summary>
    public void ShowCentered()
    {
        var screen = Screen.PrimaryScreen;
        if (screen != null)
        {
            int x = screen.WorkingArea.Width / 2;
            int y = screen.WorkingArea.Height / 2;
            Show(new System.Drawing.Point(x, y));
            Focus();
        }
        else
        {
            ShowAtCursor();
        }
    }
}
