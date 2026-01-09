using System.IO;
using System.Speech.Synthesis;
using System.Globalization;
using Microsoft.Win32;
using NAudio.Wave;
using ScreenReader.Speech;
using ScreenReader.Settings;

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
    private SpatialAudioRenderer? _spatialRenderer;
    private bool _disposed;
    private bool _isWarmedUp;

    public SpeechManager()
    {
        // Initialize SAPI5 synthesizer
        _synthesizer = new SpeechSynthesizer();
        _synthesizer.SetOutputToDefaultAudioDevice();

        // Rozgrzej syntezator SAPI5 dla lepszej responsywności
        WarmUpSynthesizer();

        // Initialize spatial renderer for 3D TTS
        _spatialRenderer = new SpatialAudioRenderer();
        _spatialRenderer.Initialize();

        // Pobierz ustawienia
        var settings = SettingsManager.Instance;

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

        // Załaduj szybkość i głośność z ustawień
        int savedRate = settings.Rate;
        int savedVolume = settings.Volume;

        _synthesizer.Rate = Math.Clamp(savedRate, -10, 10);
        _synthesizer.Volume = Math.Clamp(savedVolume, 0, 100);

        Console.WriteLine($"Załadowano ustawienia: Rate={savedRate}, Volume={savedVolume}");

        // Załaduj syntezator z ustawień
        string savedSynth = settings.Synthesizer;
        bool preferOneCore = savedSynth.Equals("OneCore", StringComparison.OrdinalIgnoreCase);

        // Initialize OneCore if available
        if (OneCoreEngine.IsAvailable())
        {
            _oneCoreEngine = new OneCoreEngine();
            if (_oneCoreEngine.Initialize())
            {
                _currentSynthesizer = preferOneCore ? SynthesizerType.OneCore : SynthesizerType.SAPI5;
                Console.WriteLine($"Syntezator z ustawień: {(preferOneCore ? "OneCore" : "SAPI5")}");

                // Zastosuj rate/volume do OneCore
                _oneCoreEngine.SetRate(savedRate);
                _oneCoreEngine.SetVolume(savedVolume);
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
            Console.WriteLine("Domyślny syntezator: SAPI5 (OneCore niedostępny)");
        }

        // Załaduj głos z ustawień
        string savedVoice = settings.Voice;
        if (!string.IsNullOrEmpty(savedVoice))
        {
            try
            {
                if (_currentSynthesizer == SynthesizerType.OneCore && _oneCoreEngine != null)
                {
                    _oneCoreEngine.SetVoice(savedVoice);
                    Console.WriteLine($"Załadowano głos OneCore: {savedVoice}");
                }
                else
                {
                    _synthesizer.SelectVoice(savedVoice);
                    Console.WriteLine($"Załadowano głos SAPI5: {savedVoice}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Nie udało się załadować zapisanego głosu: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Rozgrzewa syntezator SAPI5 dla lepszej responsywności (eliminuje opóźnienie pierwszego wywołania)
    /// </summary>
    private void WarmUpSynthesizer()
    {
        if (_isWarmedUp)
            return;

        try
        {
            // Cichy prompt do rozgrzania syntezatora - eliminuje opóźnienie pierwszego wywołania
            var prompt = new PromptBuilder();
            prompt.AppendBreak(TimeSpan.FromMilliseconds(1));
            _synthesizer.SpeakAsync(prompt);
            _synthesizer.SpeakAsyncCancelAll();
            _isWarmedUp = true;
            Console.WriteLine("SAPI5: Syntezator rozgrzany");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"SAPI5: Błąd rozgrzewania: {ex.Message}");
        }
    }

    /// <summary>
    /// Speaks text with optional 3D positioning (ONLY used during mouse exploration)
    /// For normal navigation (NumPad, keyboard), azimuth/elevation are null → normal output
    /// </summary>
    public void Speak(string text, bool interrupt = true, float? azimuth = null, float? elevation = null)
    {
        if (string.IsNullOrWhiteSpace(text))
            return;

        try
        {
            // If spatial parameters provided AND spatial renderer initialized, use 3D audio
            // This ONLY happens during mouse exploration (VirtualScreenManager)
            if (azimuth.HasValue && elevation.HasValue && _spatialRenderer?.IsInitialized == true)
            {
                SpeakSpatial(text, azimuth.Value, elevation.Value);
            }
            else
            {
                // Standard non-spatial speech for NumPad/keyboard navigation (existing code)
                if (_currentSynthesizer == SynthesizerType.SAPI5)
                {
                    if (interrupt)
                    {
                        _synthesizer.SpeakAsyncCancelAll();
                    }

                    // Użyj PromptBuilder z minimalnym czasem początku dla lepszej responsywności
                    var prompt = new PromptBuilder();
                    // Dodaj tekst bez dodatkowych przerw
                    prompt.AppendText(text);
                    _synthesizer.SpeakAsync(prompt);
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
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd mowy: {ex.Message}");
        }
    }

    /// <summary>
    /// Speaks with 3D spatial positioning
    /// </summary>
    private async void SpeakSpatial(string text, float azimuth, float elevation)
    {
        try
        {
            if (_currentSynthesizer == SynthesizerType.SAPI5)
            {
                // Capture SAPI5 to memory stream
                using var memStream = new MemoryStream();
                _synthesizer.SetOutputToWaveStream(memStream);
                _synthesizer.Speak(text);
                _synthesizer.SetOutputToDefaultAudioDevice(); // Restore

                // Decode WAV to PCM
                memStream.Position = 0;
                var pcmSamples = WavDecoder.DecodeToPCM(memStream);

                // Play through spatial audio
                _spatialRenderer?.PlaySpatial(pcmSamples, azimuth, elevation);
            }
            else if (_currentSynthesizer == SynthesizerType.OneCore && _oneCoreEngine != null)
            {
                // OneCore returns stream via SynthesizeToStreamAsync()
                var stream = await _oneCoreEngine.SynthesizeToStreamAsync(text);

                // Decode to PCM
                stream.Position = 0;
                var pcmSamples = WavDecoder.DecodeToPCM(stream);

                // Play through spatial audio
                _spatialRenderer?.PlaySpatial(pcmSamples, azimuth, elevation);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd mowy przestrzennej: {ex.Message}");
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

        _spatialRenderer?.Dispose();
        _spatialRenderer = null;

        _disposed = true;
    }
}
