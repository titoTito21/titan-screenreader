using System.Windows.Automation;

namespace ScreenReader.AppModules;

/// <summary>
/// Bazowa klasa dla modułów aplikacji
/// Port z NVDA appModuleHandler.py - AppModule
/// </summary>
public abstract class AppModuleBase : IDisposable
{
    /// <summary>Nazwa procesu (bez .exe)</summary>
    public string ProcessName { get; }

    /// <summary>ID procesu</summary>
    public int ProcessId { get; set; }

    /// <summary>Czy moduł jest aktywny</summary>
    public bool IsActive { get; protected set; }

    /// <summary>Tryb uśpienia - wyłącza czytnik dla aplikacji samodziałających</summary>
    public virtual bool SleepMode => false;

    protected AppModuleBase(string processName)
    {
        ProcessName = processName;
    }

    /// <summary>
    /// Wywoływane gdy aplikacja otrzymuje fokus
    /// </summary>
    public virtual void OnGainFocus(AutomationElement element)
    {
        IsActive = true;
    }

    /// <summary>
    /// Wywoływane gdy aplikacja traci fokus
    /// </summary>
    public virtual void OnLoseFocus()
    {
        IsActive = false;
    }

    /// <summary>
    /// Wywoływane przy zmianie fokusu wewnątrz aplikacji
    /// </summary>
    public virtual void OnFocusChanged(AutomationElement element)
    {
    }

    /// <summary>
    /// Pozwala modułowi dostosować zachowanie dla konkretnych elementów
    /// </summary>
    public virtual void CustomizeElement(AutomationElement element, ref string name, ref string role)
    {
    }

    /// <summary>
    /// Czy używać wirtualnego bufora dla tego elementu
    /// </summary>
    public virtual bool ShouldUseVirtualBuffer(AutomationElement element) => false;

    /// <summary>
    /// Pobierz tekst paska stanu aplikacji
    /// </summary>
    public virtual string? GetStatusBarText() => null;

    /// <summary>
    /// Obsługa powiadomień UIA
    /// </summary>
    public virtual void HandleNotification(string text)
    {
    }

    /// <summary>
    /// Wywoływane przy zmianie okna w aplikacji
    /// </summary>
    public virtual void OnWindowChanged(AutomationElement window)
    {
    }

    /// <summary>
    /// Filtrowanie zdarzeń zmiany właściwości UIA
    /// </summary>
    public virtual bool ShouldProcessPropertyChange(AutomationElement element, AutomationProperty property)
    {
        return true;
    }

    public virtual void Dispose()
    {
    }
}
