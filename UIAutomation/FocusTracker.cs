using System.Windows.Automation;

namespace ScreenReader;

/// <summary>
/// Śledzi zmiany fokusu w systemie Windows
/// Ulepszona wersja z lepszą obsługą błędów i stabilnością
/// </summary>
public class FocusTracker : IDisposable
{
    private AutomationFocusChangedEventHandler? _focusHandler;
    private bool _disposed;
    private bool _isRunning;
    private readonly object _lock = new();

    // Debouncing - unikaj wielokrotnych wywołań dla tego samego elementu
    private AutomationElement? _lastElement;
    private DateTime _lastFocusTime = DateTime.MinValue;
    private const int DebounceMs = 50;

    public event Action<AutomationElement>? FocusChanged;

    public void Start()
    {
        lock (_lock)
        {
            if (_isRunning || _disposed)
                return;

            try
            {
                _focusHandler = new AutomationFocusChangedEventHandler(OnFocusChanged);
                Automation.AddAutomationFocusChangedEventHandler(_focusHandler);
                _isRunning = true;
                Console.WriteLine("Focus tracking started");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error starting focus tracker: {ex.Message}");
                _focusHandler = null;
            }
        }
    }

    private void OnFocusChanged(object sender, AutomationFocusChangedEventArgs e)
    {
        if (_disposed)
            return;

        try
        {
            var element = sender as AutomationElement;
            if (element == null)
                return;

            // Debouncing - sprawdź czy to nie ten sam element
            var now = DateTime.Now;
            if ((now - _lastFocusTime).TotalMilliseconds < DebounceMs)
            {
                try
                {
                    if (_lastElement != null && Automation.Compare(element, _lastElement))
                        return;
                }
                catch
                {
                    // Ignoruj błędy porównania
                }
            }

            _lastElement = element;
            _lastFocusTime = now;

            // Sprawdź czy element jest dostępny
            try
            {
                // Próba odczytu właściwości - jeśli element jest niedostępny, wyrzuci wyjątek
                _ = element.Current.ProcessId;
            }
            catch (ElementNotAvailableException)
            {
                return; // Element zniknął, ignoruj
            }

            // Wywołaj event w bezpieczny sposób
            try
            {
                FocusChanged?.Invoke(element);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in FocusChanged handler: {ex.Message}");
            }
        }
        catch (ElementNotAvailableException)
        {
            // Element zniknął podczas przetwarzania - normalne zachowanie
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in focus changed handler: {ex.Message}");
        }
    }

    public void Stop()
    {
        lock (_lock)
        {
            if (!_isRunning)
                return;

            try
            {
                if (_focusHandler != null)
                {
                    Automation.RemoveAutomationFocusChangedEventHandler(_focusHandler);
                    _focusHandler = null;
                }
                _isRunning = false;
                _lastElement = null;
                Console.WriteLine("Focus tracking stopped");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error stopping focus tracker: {ex.Message}");
            }
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        Stop();
    }
}
