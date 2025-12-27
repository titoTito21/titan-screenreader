using System.Windows.Automation;

namespace ScreenReader.AppModules;

/// <summary>
/// Bazowy moduł dla aplikacji UWP (Windows Store Apps)
/// Zapewnia wsparcie wirtualnego bufora dla nowoczesnych aplikacji Windows
/// </summary>
public class UWPModule : AppModuleBase
{
    /// <summary>
    /// Lista znanych procesów UWP
    /// </summary>
    public static readonly string[] KnownUWPProcesses =
    {
        "ApplicationFrameHost",
        "Calculator",
        "WindowsCalculator",
        "Microsoft.WindowsCalculator",
        "SystemSettings",
        "WindowsTerminal",
        "Microsoft.WindowsTerminal",
        "Photos",
        "Microsoft.Photos",
        "Microsoft.WindowsStore",
        "WinStore.App",
        "Video.UI",
        "Microsoft.ZuneVideo",
        "Music.UI",
        "Microsoft.ZuneMusic",
        "Microsoft.WindowsAlarms",
        "Microsoft.WindowsCamera",
        "Microsoft.WindowsMaps",
        "Microsoft.GetHelp",
        "Microsoft.Getstarted",
        "Microsoft.Windows.Cortana",
        "Microsoft.WindowsFeedbackHub",
        "Microsoft.People",
        "Microsoft.MicrosoftEdge",
        "Microsoft.XboxApp",
        "Microsoft.YourPhone",
        "Microsoft.ScreenSketch",
        "Microsoft.MicrosoftStickyNotes"
    };

    public UWPModule(string processName) : base(processName)
    {
    }

    /// <summary>
    /// UWP apps often use Document, Text, and List controls that benefit from virtual buffer
    /// </summary>
    public override bool ShouldUseVirtualBuffer(AutomationElement element)
    {
        try
        {
            var controlType = element.Current.ControlType;

            // Włącz wirtualny bufor dla Document, Text i List
            return controlType == ControlType.Document ||
                   controlType == ControlType.Text ||
                   controlType == ControlType.List;
        }
        catch
        {
            return false;
        }
    }

    public override void OnGainFocus(AutomationElement element)
    {
        base.OnGainFocus(element);
        Console.WriteLine($"UWPModule: Aplikacja {ProcessName} uzyskała fokus");
    }

    public override void OnFocusChanged(AutomationElement element)
    {
        base.OnFocusChanged(element);

        try
        {
            var controlType = element.Current.ControlType;
            var name = element.Current.Name;
            var automationId = element.Current.AutomationId;

            // Log for debugging UWP element structure
            Console.WriteLine($"UWPModule: Focus -> {controlType.ProgrammaticName}, Name: {name}, AutomationId: {automationId}");
        }
        catch
        {
            // Ignore errors from stale elements
        }
    }

    public override void CustomizeElement(AutomationElement element, ref string name, ref string role)
    {
        try
        {
            // UWP apps often have generic names, try to improve them
            if (string.IsNullOrEmpty(name))
            {
                // Try to get name from child elements
                var walker = TreeWalker.ControlViewWalker;
                var child = walker.GetFirstChild(element);
                if (child != null)
                {
                    var childName = child.Current.Name;
                    if (!string.IsNullOrEmpty(childName))
                    {
                        name = childName;
                    }
                }
            }

            // Improve role descriptions for UWP controls
            var controlType = element.Current.ControlType;
            if (controlType == ControlType.Custom)
            {
                // UWP often uses Custom for specialized controls
                var localizedControlType = element.Current.LocalizedControlType;
                if (!string.IsNullOrEmpty(localizedControlType))
                {
                    role = localizedControlType;
                }
            }
        }
        catch
        {
            // Ignore errors
        }
    }

    /// <summary>
    /// Sprawdza czy proces jest znaną aplikacją UWP
    /// </summary>
    public static bool IsKnownUWPProcess(string processName)
    {
        return KnownUWPProcesses.Any(p =>
            p.Equals(processName, StringComparison.OrdinalIgnoreCase) ||
            processName.Contains(p, StringComparison.OrdinalIgnoreCase));
    }
}
