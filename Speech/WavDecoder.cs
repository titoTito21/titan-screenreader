using NAudio.Wave;

namespace ScreenReader.Speech;

/// <summary>
/// Dekoder WAV do formatu PCM dla Windows Spatial Audio
/// Używane do przechwytywania i konwersji outputu z SAPI5/OneCore TTS
/// </summary>
public static class WavDecoder
{
    /// <summary>
    /// Dekoduje strumień WAV do tablicy PCM float (mono, 48kHz)
    /// </summary>
    /// <param name="wavStream">Strumień WAV (z SAPI5 lub OneCore)</param>
    /// <returns>Tablica próbek PCM przygotowana dla Spatial Audio</returns>
    public static float[] DecodeToPCM(Stream wavStream)
    {
        using var reader = new WaveFileReader(wavStream);
        var samples = new List<float>();
        var buffer = new float[4096];
        int count;

        // Konwertuj do ISampleProvider
        var provider = reader.ToSampleProvider();

        // Konwertuj stereo → mono jeśli potrzebne
        if (provider.WaveFormat.Channels > 1)
        {
            provider = provider.ToMono();
            Console.WriteLine("WavDecoder: Konwersja stereo → mono");
        }

        // Resample do 48kHz jeśli potrzebne (wymagane przez Spatial Audio)
        if (provider.WaveFormat.SampleRate != 48000)
        {
            Console.WriteLine($"WavDecoder: Resampling {provider.WaveFormat.SampleRate}Hz → 48000Hz");
            var targetFormat = WaveFormat.CreateIeeeFloatWaveFormat(48000, provider.WaveFormat.Channels);
            var resampler = new MediaFoundationResampler(provider.ToWaveProvider(), targetFormat);
            provider = resampler.ToSampleProvider();
        }

        // Odczytaj wszystkie próbki
        while ((count = provider.Read(buffer, 0, buffer.Length)) > 0)
        {
            for (int i = 0; i < count; i++)
                samples.Add(buffer[i]);
        }

        Console.WriteLine($"WavDecoder: Zdekodowano {samples.Count} próbek ({samples.Count / 48000.0:F2}s @ 48kHz)");
        return samples.ToArray();
    }
}
