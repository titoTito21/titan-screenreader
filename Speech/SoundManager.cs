using NAudio.Wave;
using NAudio.Vorbis;

namespace ScreenReader;

public class SoundManager : IDisposable
{
    private readonly string _soundsDirectory;
    private bool _disposed;

    public SoundManager(string soundsDirectory)
    {
        _soundsDirectory = soundsDirectory;
    }

    public void PlaySound(string soundFileName, float pitch = 1.0f)
    {
        try
        {
            var soundPath = Path.Combine(_soundsDirectory, soundFileName);
            
            if (!File.Exists(soundPath))
            {
                Console.WriteLine($"Brak pliku dźwiękowego: {soundPath}");
                return;
            }

            // Don't stop previous sounds - allow overlap with TTS
            var waveOut = new WaveOutEvent();
            
            var vorbisReader = new VorbisWaveReader(soundPath);
            
            // Apply pitch shifting if needed
            ISampleProvider sampleProvider = vorbisReader.ToSampleProvider();
            if (Math.Abs(pitch - 1.0f) > 0.01f)
            {
                sampleProvider = new PitchShiftingSampleProvider(sampleProvider, pitch);
            }

            waveOut.Init(sampleProvider);
            
            // Dispose after playback finishes
            waveOut.PlaybackStopped += (s, e) =>
            {
                waveOut.Dispose();
                vorbisReader.Dispose();
            };
            
            waveOut.Play();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Błąd odtwarzania dźwięku: {ex.Message}");
        }
    }

    public void PlayClicked()
    {
        PlaySound("clicked.ogg");
    }

    public void PlayCursor()
    {
        PlaySound("cursor.ogg");
    }

    public void PlayEdge()
    {
        PlaySound("edge.ogg");
    }

    public void PlayListItem(float position)
    {
        // Position: 0.0 = top (higher pitch), 1.0 = bottom (lower pitch)
        // Map to pitch range: 1.5 (high) to 0.7 (low)
        float pitch = 1.5f - (position * 0.8f);
        PlaySound("listitem.ogg", pitch);
    }

    public void PlayWindow()
    {
        PlaySound("window.ogg");
    }

    public void Stop()
    {
        // With overlapping sounds, Stop() doesn't need to do anything
        // Each sound will dispose itself when finished
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
    }
}

// Simple pitch shifting by changing playback rate
public class PitchShiftingSampleProvider : ISampleProvider
{
    private readonly ISampleProvider _source;
    private readonly float _pitchFactor;

    public PitchShiftingSampleProvider(ISampleProvider source, float pitchFactor)
    {
        _source = source;
        _pitchFactor = pitchFactor;
        WaveFormat = WaveFormat.CreateIeeeFloatWaveFormat(
            (int)(source.WaveFormat.SampleRate * pitchFactor),
            source.WaveFormat.Channels);
    }

    public WaveFormat WaveFormat { get; }

    public int Read(float[] buffer, int offset, int count)
    {
        return _source.Read(buffer, offset, count);
    }
}
