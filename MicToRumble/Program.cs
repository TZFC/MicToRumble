using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using SharpDX.XInput;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new Mic_to_Rumble_Form());
    }
}

public class Mic_to_Rumble_Form : Form
{
    private ComboBox input_device_combo_box;
    private Button start_button;
    private Button stop_button;
    private TrackBar threshold_track_bar_dbfs;
    private NumericUpDown threshold_numeric_up_down_dbfs;
    private Label current_level_label_dbfs;
    private ProgressBar level_progress_bar_linear;
    private Button test_rumble_button;
    private Label controller_status_label;

    private MMDeviceEnumerator mm_device_enumerator;
    private WasapiCapture wasapi_capture;
    private float current_rms_linear;
    private float current_level_dbfs;
    private float threshold_dbfs = -25.0f;

    private Controller xinput_controller = new Controller(UserIndex.One);
    private System.Windows.Forms.Timer ui_update_timer;

    private int capture_sample_rate_hz = 48000;
    private int capture_buffer_milliseconds = 50;

    private ushort rumble_strength_left_motor = 20000;   // 0..65535
    private ushort rumble_strength_right_motor = 20000;  // 0..65535
    private int rumble_duration_milliseconds = 150;
    private DateTime last_rumble_end_time_utc = DateTime.MinValue;

    public Mic_to_Rumble_Form()
    {
        Text = "Microphone → Xbox Controller Rumble";
        Size = new Size(580, 300);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        mm_device_enumerator = new MMDeviceEnumerator();

        input_device_combo_box = new ComboBox
        {
            Left = 16, Top = 16, Width = 360, DropDownStyle = ComboBoxStyle.DropDownList
        };
        start_button = new Button { Left = 392, Top = 16, Width = 80, Text = "Start" };
        stop_button = new Button { Left = 478, Top = 16, Width = 80, Text = "Stop", Enabled = false };

        var threshold_label = new Label { Left = 16, Top = 60, Width = 220, Text = "Threshold (decibels full scale):" };
        threshold_track_bar_dbfs = new TrackBar
        {
            Left = 16, Top = 82, Width = 350,
            Minimum = -60, Maximum = 0, TickFrequency = 5, Value = (int)Math.Round(threshold_dbfs)
        };
        threshold_numeric_up_down_dbfs = new NumericUpDown
        {
            Left = 380, Top = 82, Width = 80,
            Minimum = -120, Maximum = 0, DecimalPlaces = 0, Value = (decimal)threshold_dbfs
        };

        var current_level_text_label = new Label { Left = 16, Top = 130, Width = 120, Text = "Current Level:" };
        current_level_label_dbfs = new Label { Left = 132, Top = 130, Width = 120, Text = "-∞ dBFS" };
        level_progress_bar_linear = new ProgressBar { Left = 16, Top = 152, Width = 540, Height = 20, Minimum = 0, Maximum = 6000 };

        test_rumble_button = new Button { Left = 16, Top = 186, Width = 140, Text = "Test Rumble" };
        controller_status_label = new Label { Left = 172, Top = 190, Width = 384, Text = "Controller: Unknown" };

        Controls.Add(input_device_combo_box);
        Controls.Add(start_button);
        Controls.Add(stop_button);
        Controls.Add(threshold_label);
        Controls.Add(threshold_track_bar_dbfs);
        Controls.Add(threshold_numeric_up_down_dbfs);
        Controls.Add(current_level_text_label);
        Controls.Add(current_level_label_dbfs);
        Controls.Add(level_progress_bar_linear);
        Controls.Add(test_rumble_button);
        Controls.Add(controller_status_label);

        Load += on_form_load;
        FormClosing += on_form_closing;
        start_button.Click += on_start_clicked;
        stop_button.Click += on_stop_clicked;
        test_rumble_button.Click += on_test_rumble_clicked;

        threshold_track_bar_dbfs.Scroll += (s, e) =>
        {
            threshold_numeric_up_down_dbfs.Value = threshold_track_bar_dbfs.Value;
            threshold_dbfs = threshold_track_bar_dbfs.Value;
        };
        threshold_numeric_up_down_dbfs.ValueChanged += (s, e) =>
        {
            threshold_track_bar_dbfs.Value = (int)threshold_numeric_up_down_dbfs.Value;
            threshold_dbfs = (float)threshold_numeric_up_down_dbfs.Value;
        };

        ui_update_timer = new System.Windows.Forms.Timer();
        ui_update_timer.Interval = 50;
        ui_update_timer.Tick += (s, e) => update_ui_from_meter();
        ui_update_timer.Start();
    }

    private void on_form_load(object sender, EventArgs e)
    {
        populate_input_devices();
        update_controller_status_label();
    }

    private void on_form_closing(object sender, FormClosingEventArgs e)
    {
        stop_capture_if_active();
        try { set_rumble(0, 0); } catch { }
        mm_device_enumerator?.Dispose();
    }

    private void on_start_clicked(object sender, EventArgs e)
    {
        if (wasapi_capture != null) return;

        var device_item = input_device_combo_box.SelectedItem as Recording_device_item;
        if (device_item == null)
        {
            MessageBox.Show("Please select a recording device.", "No Device", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!xinput_controller.IsConnected)
        {
            MessageBox.Show("No XInput controller detected. Please connect an Xbox controller and try again.", "Controller Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            update_controller_status_label();
            return;
        }

        try
        {
            wasapi_capture = new WasapiCapture(device_item.mm_device, false, capture_buffer_milliseconds)
            {
                ShareMode = AudioClientShareMode.Shared,
                WaveFormat = new WaveFormat(capture_sample_rate_hz, 16, 1)
            };
            wasapi_capture.DataAvailable += on_audio_data_available;
            wasapi_capture.StartRecording();

            start_button.Enabled = false;
            stop_button.Enabled = true;
            input_device_combo_box.Enabled = false;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Failed to start capture: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            stop_capture_if_active();
        }
    }

    private void on_stop_clicked(object sender, EventArgs e)
    {
        stop_capture_if_active();
    }

    private void on_test_rumble_clicked(object sender, EventArgs e)
    {
        if (!xinput_controller.IsConnected)
        {
            MessageBox.Show("No XInput controller detected.", "Controller Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            update_controller_status_label();
            return;
        }
        trigger_rumble_once();
    }

    private void populate_input_devices()
    {
        input_device_combo_box.Items.Clear();
        var devices = mm_device_enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active);
        foreach (var device in devices)
        {
            input_device_combo_box.Items.Add(new Recording_device_item(device));
        }
        if (input_device_combo_box.Items.Count > 0)
        {
            input_device_combo_box.SelectedIndex = 0;
        }
    }

    private void update_controller_status_label()
    {
        controller_status_label.Text = xinput_controller.IsConnected
            ? "Controller: Connected (UserIndex.One)"
            : "Controller: Not Connected";
    }

    private void stop_capture_if_active()
    {
        if (wasapi_capture != null)
        {
            try { wasapi_capture.StopRecording(); } catch { }
            try { wasapi_capture.Dispose(); } catch { }
            wasapi_capture = null;
        }

        start_button.Enabled = true;
        stop_button.Enabled = false;
        input_device_combo_box.Enabled = true;
    }

    private void on_audio_data_available(object sender, WaveInEventArgs e)
    {
        int sample_count = e.BytesRecorded / 2;
        if (sample_count <= 0) return;

        double sum_of_squares = 0.0;
        for (int i = 0; i < sample_count; i++)
        {
            short sample_16_bit = BitConverter.ToInt16(e.Buffer, i * 2);
            float normalized_linear = sample_16_bit / 32768.0f; // -1.0 to +1.0
            sum_of_squares += normalized_linear * normalized_linear;
        }

        float mean_square = (float)(sum_of_squares / sample_count);
        float rms_linear = (float)Math.Sqrt(mean_square);
        float dbfs = (rms_linear > 0f) ? 20.0f * (float)Math.Log10(rms_linear) : -120.0f;

        current_rms_linear = rms_linear;
        current_level_dbfs = dbfs;

        if (current_level_dbfs >= threshold_dbfs)
        {
            trigger_rumble_once();
        }
    }

    private void update_ui_from_meter()
    {
        var displayed_dbfs = float.IsFinite(current_level_dbfs) ? current_level_dbfs : -120.0f;
        current_level_label_dbfs.Text = $"{displayed_dbfs:0} dBFS";

        int bar_value = (int)Math.Max(0, Math.Min(6000, (displayed_dbfs + 120f) / 120f * 6000f));
        level_progress_bar_linear.Value = bar_value;
    }

    private void trigger_rumble_once()
    {
        if (DateTime.UtcNow < last_rumble_end_time_utc)
        {
            return;
        }

        set_rumble(rumble_strength_left_motor, rumble_strength_right_motor);
        last_rumble_end_time_utc = DateTime.UtcNow.AddMilliseconds(rumble_duration_milliseconds);

        new Thread(() =>
        {
            Thread.Sleep(rumble_duration_milliseconds);
            set_rumble(0, 0);
        })
        { IsBackground = true }.Start();
    }

    private void set_rumble(ushort left_motor_speed, ushort right_motor_speed)
    {
        var vibration = new Vibration
        {
            LeftMotorSpeed = left_motor_speed,
            RightMotorSpeed = right_motor_speed
        };
        xinput_controller.SetVibration(vibration);
    }

    private class Recording_device_item
    {
        public MMDevice mm_device;

        public Recording_device_item(MMDevice device)
        {
            mm_device = device;
        }

        public override string ToString()
        {
            try { return $"{mm_device.FriendlyName}"; }
            catch { return "Unknown Device"; }
        }
    }
}
