using System.Windows.Forms;
using System.Speech.Synthesis;
using ScreenReader.Speech;

namespace ScreenReader;

public class SettingsDialog : Form
{
    private readonly SpeechManager _speechManager;
    private ComboBox _synthesizerComboBox = null!;
    private ComboBox _voiceComboBox = null!;
    private Label _voiceLabel = null!;
    private NumericUpDown _rateNumeric = null!;
    private NumericUpDown _volumeNumeric = null!;
    private NumericUpDown _pitchNumeric = null!;

    // Przechowuje mapowanie wyświetlanej nazwy na identyfikator głosu OneCore
    private Dictionary<string, string> _oneCoreVoiceMap = new();

    public SettingsDialog(SpeechManager speechManager)
    {
        _speechManager = speechManager;
        InitializeComponents();
        LoadCurrentSettings();
    }

    private void InitializeComponents()
    {
        Text = "Ustawienia - Tekst na mowę";
        Width = 450;
        Height = 350;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        TopMost = true;

        int yPos = 20;

        // Synthesizer selection
        var lblSynthesizer = new Label
        {
            Text = "Syntezator:",
            Left = 20,
            Top = yPos,
            Width = 100
        };
        _synthesizerComboBox = new ComboBox
        {
            Left = 130,
            Top = yPos,
            Width = 280,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _synthesizerComboBox.Items.Add("SAPI 5");
        _synthesizerComboBox.Items.Add("OneCore (Mobile)");
        _synthesizerComboBox.SelectedIndexChanged += SynthesizerComboBox_SelectedIndexChanged;
        yPos += 40;

        // Voice selection
        _voiceLabel = new Label
        {
            Text = "Głos:",
            Left = 20,
            Top = yPos,
            Width = 100
        };
        _voiceComboBox = new ComboBox
        {
            Left = 130,
            Top = yPos,
            Width = 280,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _voiceComboBox.SelectedIndexChanged += VoiceComboBox_SelectedIndexChanged;
        yPos += 40;

        // Rate
        var lblRate = new Label
        {
            Text = "Szybkość:",
            Left = 20,
            Top = yPos,
            Width = 100
        };
        _rateNumeric = new NumericUpDown
        {
            Left = 130,
            Top = yPos,
            Width = 100,
            Minimum = -10,
            Maximum = 10,
            Value = 0
        };
        var lblRateInfo = new Label
        {
            Text = "(-10 wolno, 0 normalnie, 10 szybko)",
            Left = 240,
            Top = yPos + 3,
            Width = 200,
            AutoSize = true
        };
        yPos += 40;

        // Volume
        var lblVolume = new Label
        {
            Text = "Głośność:",
            Left = 20,
            Top = yPos,
            Width = 100
        };
        _volumeNumeric = new NumericUpDown
        {
            Left = 130,
            Top = yPos,
            Width = 100,
            Minimum = 0,
            Maximum = 100,
            Value = 100
        };
        var lblVolumeInfo = new Label
        {
            Text = "(0-100)",
            Left = 240,
            Top = yPos + 3,
            Width = 100,
            AutoSize = true
        };
        yPos += 40;

        // Pitch (note: may not work with all voices)
        var lblPitch = new Label
        {
            Text = "Wysokość:",
            Left = 20,
            Top = yPos,
            Width = 100
        };
        _pitchNumeric = new NumericUpDown
        {
            Left = 130,
            Top = yPos,
            Width = 100,
            Minimum = -10,
            Maximum = 10,
            Value = 0,
            Enabled = false // Disabled for now - System.Speech doesn't support pitch
        };
        var lblPitchInfo = new Label
        {
            Text = "(nieobsługiwane przez System.Speech)",
            Left = 240,
            Top = yPos + 3,
            Width = 200,
            AutoSize = true,
            ForeColor = System.Drawing.Color.Gray
        };
        yPos += 60;

        // Buttons
        var btnOK = new Button
        {
            Text = "OK",
            Left = 150,
            Top = yPos,
            Width = 80
        };
        btnOK.Click += (s, e) =>
        {
            ApplySettings();
            DialogResult = DialogResult.OK;
            Close();
        };

        var btnCancel = new Button
        {
            Text = "Anuluj",
            Left = 240,
            Top = yPos,
            Width = 80,
            DialogResult = DialogResult.Cancel
        };

        var btnApply = new Button
        {
            Text = "Zastosuj",
            Left = 330,
            Top = yPos,
            Width = 80
        };
        btnApply.Click += (s, e) => ApplySettings();

        Controls.AddRange(new Control[] {
            lblSynthesizer, _synthesizerComboBox,
            _voiceLabel, _voiceComboBox,
            lblRate, _rateNumeric, lblRateInfo,
            lblVolume, _volumeNumeric, lblVolumeInfo,
            lblPitch, _pitchNumeric, lblPitchInfo,
            btnOK, btnCancel, btnApply
        });

        AcceptButton = btnOK;
        CancelButton = btnCancel;
    }

    private void SynthesizerComboBox_SelectedIndexChanged(object? sender, EventArgs e)
    {
        // Update voice list based on synthesizer
        UpdateVoiceList();
    }

    private void VoiceComboBox_SelectedIndexChanged(object? sender, EventArgs e)
    {
        // No need to update variants - they're combined in the voice list
    }

    private void UpdateVoiceList()
    {
        _voiceComboBox.Items.Clear();
        _oneCoreVoiceMap.Clear();

        bool isOneCore = _synthesizerComboBox.SelectedIndex == 1;

        if (isOneCore)
        {
            // Load OneCore voices with full information
            var voices = _speechManager.GetOneCoreVoicesInfo();
            int plIndex = -1;
            int index = 0;

            foreach (var voice in voices)
            {
                // Formatuj wyświetlaną nazwę: "Microsoft Paulina (pl-PL) [F]"
                string displayName = voice.ToString();
                _voiceComboBox.Items.Add(displayName);
                _oneCoreVoiceMap[displayName] = voice.Id;

                // Znajdź pierwszy polski głos
                if (plIndex < 0 && voice.Language.StartsWith("pl"))
                {
                    plIndex = index;
                }
                index++;
            }

            // Wybierz polski głos lub pierwszy dostępny
            if (_voiceComboBox.Items.Count > 0)
            {
                _voiceComboBox.SelectedIndex = plIndex >= 0 ? plIndex : 0;
            }
        }
        else
        {
            // Load SAPI5 voices
            var voices = _speechManager.GetAvailableVoices();
            foreach (var voice in voices)
            {
                _voiceComboBox.Items.Add(voice);
            }

            // Select current voice
            var currentVoice = _speechManager.GetCurrentVoice();
            if (!string.IsNullOrEmpty(currentVoice))
            {
                int index = _voiceComboBox.Items.IndexOf(currentVoice);
                if (index >= 0)
                    _voiceComboBox.SelectedIndex = index;
            }
            else if (_voiceComboBox.Items.Count > 0)
            {
                _voiceComboBox.SelectedIndex = 0;
            }
        }
    }

    private void LoadCurrentSettings()
    {
        // Select current synthesizer
        var currentSynth = _speechManager.GetCurrentSynthesizer();
        _synthesizerComboBox.SelectedIndex = currentSynth == SynthesizerType.SAPI5 ? 0 : 1;

        // Voice list will be populated by the SelectedIndexChanged event

        // Load current settings
        _rateNumeric.Value = _speechManager.GetRate();
        _volumeNumeric.Value = _speechManager.GetVolume();
    }

    private void ApplySettings()
    {
        // Apply synthesizer
        bool isOneCore = _synthesizerComboBox.SelectedIndex == 1;
        _speechManager.SetSynthesizer(isOneCore ? SynthesizerType.OneCore : SynthesizerType.SAPI5);

        // Apply voice
        if (_voiceComboBox.SelectedItem != null)
        {
            string selectedItem = _voiceComboBox.SelectedItem.ToString()!;

            if (isOneCore)
            {
                // Użyj mapy do uzyskania identyfikatora głosu
                if (_oneCoreVoiceMap.TryGetValue(selectedItem, out string? voiceId))
                {
                    _speechManager.SetOneCoreVoice(voiceId);
                }
                else
                {
                    // Fallback - użyj wybranego tekstu bezpośrednio
                    _speechManager.SetOneCoreVoice(selectedItem);
                }
            }
            else
            {
                _speechManager.SelectVoice(selectedItem);
            }
        }

        // Apply rate and volume
        _speechManager.SetRate((int)_rateNumeric.Value);
        _speechManager.SetVolume((int)_volumeNumeric.Value);

        // Test speech
        _speechManager.Speak("Ustawienia zastosowane");
    }
}
