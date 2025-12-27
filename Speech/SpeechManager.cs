using System.Speech.Synthesis;
using System.Globalization;
using Microsoft.Win32;
using ScreenReader.Speech;

namespace ScreenReader;

public enum SynthesizerType
{
    SAPI5,
    OneCore
}

public class SpeechManager : IDisposable
{
    private readonly SpeechSynthesizer _synthesizer;
    private OneCoreEngine? _oneCoreEngine;
    private SynthesizerType _currentSynthesizer;
    private bool _disposed;

    public SpeechManager()
    {
        // Initialize SAPI5 synthesizer
        _synthesizer = new SpeechSynthesizer();
        _synthesizer.SetOutputToDefaultAudioDevice();

        // Try to set Polish voice
        try
        {
            var polishVoice = _synthesizer.GetInstalledVoices()
                .FirstOrDefault(v => v.VoiceInfo.Culture.TwoLetterISOLanguageName == "pl");

            if (polishVoice != null)
            {
                _synthesizer.SelectVoice(polishVoice.VoiceInfo.Name);
                Console.WriteLine($"Używam głosu SAPI5: {polishVoice.VoiceInfo.Name}");
            }
            else
            {
                Console.WriteLine("Uwaga: Brak polskiego głosu TTS. Używam domyślnego.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd wyboru głosu: {ex.Message}");
        }

        // Set default rate (normal speed)
        _synthesizer.Rate = 0;

        // Set default volume
        _synthesizer.Volume = 100;

        // Default to OneCore if available, otherwise SAPI5
        if (OneCoreEngine.IsAvailable())
        {
            _oneCoreEngine = new OneCoreEngine();
            if (_oneCoreEngine.Initialize())
            {
                _currentSynthesizer = SynthesizerType.OneCore;
                Console.WriteLine("Domyślny syntezator: OneCore");
            }
            else
            {
                _currentSynthesizer = SynthesizerType.SAPI5;
                _oneCoreEngine?.Dispose();
                _oneCoreEngine = null;
            }
        }
        else
        {
            _currentSynthesizer = SynthesizerType.SAPI5;
            Console.WriteLine("Domyślny syntezator: SAPI5");
        }
    }

    public void Speak(string text, bool interrupt = true)
    {
        if (string.IsNullOrWhiteSpace(text))
            return;

        try
        {
            if (_currentSynthesizer == SynthesizerType.SAPI5)
            {
                if (interrupt)
                {
                    _synthesizer.SpeakAsyncCancelAll();
                }
                _synthesizer.SpeakAsync(text);
            }
            else if (_currentSynthesizer == SynthesizerType.OneCore)
            {
                if (_oneCoreEngine != null)
                {
                    // OneCore zawsze przerywa poprzednią mowę
                    _oneCoreEngine.Speak(text);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd mowy: {ex.Message}");
        }
    }

    public void Stop()
    {
        if (_currentSynthesizer == SynthesizerType.SAPI5)
        {
            _synthesizer.SpeakAsyncCancelAll();
        }
        else if (_oneCoreEngine != null)
        {
            _oneCoreEngine.Stop();
        }
    }

    public void SetRate(int rate)
    {
        // Rate range: -10 (slow) to 10 (fast)
        if (_currentSynthesizer == SynthesizerType.SAPI5)
        {
            _synthesizer.Rate = Math.Clamp(rate, -10, 10);
        }
        else if (_oneCoreEngine != null)
        {
            _oneCoreEngine.SetRate(rate);
        }
    }

    public void SetVolume(int volume)
    {
        // Volume range: 0 to 100
        if (_currentSynthesizer == SynthesizerType.SAPI5)
        {
            _synthesizer.Volume = Math.Clamp(volume, 0, 100);
        }
        else if (_oneCoreEngine != null)
        {
            _oneCoreEngine.SetVolume(volume);
        }
    }

    public List<string> GetAvailableVoices()
    {
        if (_currentSynthesizer == SynthesizerType.SAPI5)
        {
            // Użyj GetInstalledVoices() - zwraca prawidłowe nazwy dla SelectVoice()
            return _synthesizer.GetInstalledVoices()
                .Where(v => v.Enabled)
                .Select(v => v.VoiceInfo.Name)
                .ToList();
        }
        else
        {
            return OneCoreEngine.GetAllVoices()
                .Select(v => v.DisplayName)
                .ToList();
        }
    }

    public string GetCurrentVoice()
    {
        if (_currentSynthesizer == SynthesizerType.SAPI5)
        {
            return _synthesizer.Voice.Name;
        }
        else if (_oneCoreEngine != null)
        {
            return _oneCoreEngine.GetCurrentVoice();
        }
        return "";
    }

    public void SelectVoice(string voiceName)
    {
        try
        {
            if (_currentSynthesizer == SynthesizerType.SAPI5)
            {
                _synthesizer.SelectVoice(voiceName);
                Console.WriteLine($"Zmieniono głos SAPI5 na: {voiceName}");
            }
            else if (_oneCoreEngine != null)
            {
                _oneCoreEngine.SetVoice(voiceName);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd zmiany głosu: {ex.Message}");
        }
    }

    public int GetRate()
    {
        return _synthesizer.Rate;
    }

    public int GetVolume()
    {
        return _synthesizer.Volume;
    }

    // SAPI voice detection including 32-bit voices
    private List<string> GetAllSAPIVoices()
    {
        var voices = new HashSet<string>();

        // 64-bitowe głosy SAPI
        try
        {
            using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Speech\Voices\Tokens"))
            {
                if (key != null)
                {
                    foreach (var tokenName in key.GetSubKeyNames())
                    {
                        using (var voiceKey = key.OpenSubKey(tokenName))
                        {
                            if (voiceKey != null)
                            {
                                // Odczytaj nazwę głosu z wartości domyślnej
                                var voiceName = voiceKey.GetValue("") as string;
                                if (!string.IsNullOrEmpty(voiceName))
                                {
                                    voices.Add(voiceName);
                                }
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd odczytu głosów 64-bit: {ex.Message}");
        }

        // 32-bitowe głosy SAPI (Wow6432Node)
        try
        {
            using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Speech\Voices\Tokens"))
            {
                if (key != null)
                {
                    foreach (var tokenName in key.GetSubKeyNames())
                    {
                        using (var voiceKey = key.OpenSubKey(tokenName))
                        {
                            if (voiceKey != null)
                            {
                                // Odczytaj nazwę głosu z wartości domyślnej
                                var voiceName = voiceKey.GetValue("") as string;
                                if (!string.IsNullOrEmpty(voiceName))
                                {
                                    voices.Add(voiceName);
                                }
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd odczytu głosów 32-bit: {ex.Message}");
        }

        return voices.OrderBy(v => v).ToList();
    }

    // Synthesizer management
    public SynthesizerType GetCurrentSynthesizer()
    {
        return _currentSynthesizer;
    }

    public void SetSynthesizer(SynthesizerType type)
    {
        if (_currentSynthesizer == type)
            return;

        _currentSynthesizer = type;

        if (type == SynthesizerType.OneCore)
        {
            // Initialize OneCore if not already done
            if (_oneCoreEngine == null)
            {
                _oneCoreEngine = new OneCoreEngine();
                if (!_oneCoreEngine.Initialize())
                {
                    Console.WriteLine("Nie udało się zainicjalizować OneCore, powrót do SAPI5");
                    _currentSynthesizer = SynthesizerType.SAPI5;
                    _oneCoreEngine?.Dispose();
                    _oneCoreEngine = null;
                }
            }
        }
    }

    // OneCore-specific methods
    public void SetOneCoreVoice(string voiceId)
    {
        if (_currentSynthesizer == SynthesizerType.OneCore && _oneCoreEngine != null)
        {
            _oneCoreEngine.SetVoice(voiceId);
        }
    }

    public List<OneCoreEngine.VoiceInfo> GetOneCoreVoicesInfo()
    {
        return OneCoreEngine.GetAllVoices();
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _synthesizer.SpeakAsyncCancelAll();
        _synthesizer.Dispose();

        _oneCoreEngine?.Dispose();
        _oneCoreEngine = null;

        _disposed = true;
    }
}
