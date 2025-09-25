// Program.cs
// Single-file WinForms app that listens to a selected microphone,
// applies attack–release smoothing, uses dwell (over-threshold hold)
// and cooldown to prevent spamming, and vibrates an Xbox controller
// when the smoothed level exceeds a threshold.
//
// Dependencies (run in project folder):
// dotnet add package NAudio
// dotnet add package SharpDX.XInput
//
// Build and run:
// dotnet new winforms -n MicToRumble
// cd MicToRumble
// dotnet add package NAudio
// dotnet add package SharpDX.XInput
// Replace Program.cs with this file, then:
// dotnet run

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
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new mic_to_rumble_form());
    }
}

public class mic_to_rumble_form : Form
{
    // -------------------- UI CONTROLS --------------------
    private ComboBox input_device_combo_box;
    private Button start_button;
    private Button stop_button;

    private Label threshold_label;
    private TrackBar threshold_trackbar; // -80..0 dBFS
    private Label threshold_value_label;

    private Label attack_time_label;
    private TrackBar attack_time_trackbar; // 1..200 ms
    private Label attack_time_value_label;

    private Label release_time_label;
    private TrackBar release_time_trackbar; // 10..800 ms
    private Label release_time_value_label;

    private Label dwell_time_label;
    private TrackBar dwell_time_trackbar; // 0..1000 ms
    private Label dwell_time_value_label;

    private Label cooldown_time_label;
    private TrackBar cooldown_time_trackbar; // 0..2000 ms
    private Label cooldown_time_value_label;

    private Label current_level_label;
    private ProgressBar current_level_progress_bar;

    private Button test_rumble_button;
    private Label controller_status_label;

    private Label rumble_time_label;
    private TrackBar rumble_time_trackbar;
    private Label rumble_time_value_label;

    // -------------------- AUDIO AND CONTROLLER --------------------
    private MMDeviceEnumerator audio_device_enumerator;
    private WasapiCapture wasapi_capture;
    private Controller xinput_controller = new Controller(UserIndex.One);

    // -------------------- ENVELOPE AND TRIGGER STATE --------------------
    private float current_level_rms_linear = 0.0f;
    private float smoothed_level_rms_linear = 0.0f;
    private float current_level_dbfs = -120.0f;

    private float threshold_dbfs = -10.0f;
    private int attack_time_milliseconds = 15;
    private int release_time_milliseconds = 150;
    private int dwell_time_milliseconds = 80;
    private int cooldown_time_milliseconds = 250;

    private DateTime last_buffer_time_utc = DateTime.UtcNow;
    private int consecutive_over_threshold_time_milliseconds = 0;
    private DateTime last_rumble_end_time_utc = DateTime.MinValue;

    // -------------------- CAPTURE AND RUMBLE --------------------
    private readonly int capture_buffer_milliseconds = 50;

    private ushort rumble_strength_left_motor = 65535;   // 0..65535
    private ushort rumble_strength_right_motor = 65535;  // 0..65535
    private int rumble_duration_milliseconds = 1000;

    // -------------------- UI TIMER --------------------
    private System.Windows.Forms.Timer ui_update_timer;

    public mic_to_rumble_form()
    {
        // basic window styling for a cleaner, modern-ish look
        Text = "Microphone → Xbox Controller Rumble";
        MinimumSize = new Size(960, 640);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

        // top panel with device selector and start/stop
        var top_panel = new Panel { Dock = DockStyle.Top, Height = 70, Padding = new Padding(16) };
        var device_label = new Label { Text = "Recording device:", AutoSize = true, Left = 0, Top = 8 };
        input_device_combo_box = new ComboBox
        {
            Left = 0,
            Top = 28,
            Width = 520,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        start_button = new Button { Text = "Start", Width = 100, Left = 540, Top = 26, Height = 40};
        stop_button = new Button { Text = "Stop", Width = 100, Left = 650, Top = 26, Height = 40, Enabled = false };
        test_rumble_button = new Button { Text = "Test Rumble", Width = 120, Left = 760, Top = 26, Height = 40 };

        top_panel.Controls.Add(device_label);
        top_panel.Controls.Add(input_device_combo_box);
        top_panel.Controls.Add(start_button);
        top_panel.Controls.Add(stop_button);
        top_panel.Controls.Add(test_rumble_button);

        // left panel with sliders
        var left_panel = new Panel { Dock = DockStyle.Left, Width = 560, Padding = new Padding(16), AutoScroll = true };

        var threshold_group = make_group_box("Threshold (dBFS)", out threshold_trackbar, out threshold_value_label, -40, 20, -10, 1, " dBFS");
        var attack_group    = make_group_box("Attack time (milliseconds)", out attack_time_trackbar, out attack_time_value_label, 1, 200, 15, 1, " ms");
        var release_group   = make_group_box("Release time (milliseconds)", out release_time_trackbar, out release_time_value_label, 10, 800, 150, 10, " ms");
        var dwell_group     = make_group_box("Over-threshold hold (dwell) (milliseconds)", out dwell_time_trackbar, out dwell_time_value_label, 0, 1000, 80, 10, " ms");
        var cooldown_group  = make_group_box("Cooldown between rumbles (milliseconds)", out cooldown_time_trackbar, out cooldown_time_value_label, 0, 2000, 250, 20, " ms");
        var rumble_group    = make_group_box("Minimum rumble time (milliseconds)", out rumble_time_trackbar, out rumble_time_value_label, 500, 10000, 1000, 50, " ms");

        var left_layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 6,
            Padding = new Padding(0),
            AutoSize = false,      // IMPORTANT: do not autosize the table
            AutoScroll = true,     // allow scrolling if needed
            GrowStyle = TableLayoutPanelGrowStyle.FixedSize
        };
        left_layout.ColumnStyles.Clear();
        left_layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        left_layout.RowStyles.Clear();
        // Give each row a fixed (absolute) height so it never collapses:
        for (int i = 0; i < 6; i++) left_layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 120f));

        left_layout.Controls.Add(threshold_group, 0, 0);
        left_layout.Controls.Add(attack_group,    0, 1);
        left_layout.Controls.Add(release_group,   0, 2);
        left_layout.Controls.Add(dwell_group,     0, 3);
        left_layout.Controls.Add(cooldown_group,  0, 4);
        left_layout.Controls.Add(rumble_group,    0, 5);

        left_panel.Controls.Add(left_layout);


        // right panel with meter and controller status
        var right_panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(16) };
        var meter_group = new GroupBox { Text = "Input level", Dock = DockStyle.Top, Height = 140, Padding = new Padding(16) };
        current_level_label = new Label { Text = "Current level: -∞ dBFS", AutoSize = true, Left = 8, Top = 28 };
        current_level_progress_bar = new ProgressBar { Left = 8, Top = 56, Width = 320, Height = 24, Minimum = 0, Maximum = 100, Value = 0 };
        meter_group.Controls.Add(current_level_label);
        meter_group.Controls.Add(current_level_progress_bar);

        var controller_group = new GroupBox { Text = "Controller", Dock = DockStyle.Top, Height = 90, Padding = new Padding(16) };
        controller_status_label = new Label { Text = "Controller: Unknown", AutoSize = true, Left = 8, Top = 28 };
        controller_group.Controls.Add(controller_status_label);

        right_panel.Controls.Add(controller_group);
        right_panel.Controls.Add(meter_group);

        Controls.Add(right_panel);
        Controls.Add(left_panel);
        Controls.Add(top_panel);

        // event wiring
        Load += on_form_load;
        FormClosing += on_form_closing;

        start_button.Click += on_start_clicked;
        stop_button.Click += on_stop_clicked;
        test_rumble_button.Click += on_test_rumble_clicked;

        rumble_time_trackbar.Scroll += (_, __) =>
        {
            rumble_duration_milliseconds = rumble_time_trackbar.Value;
            rumble_time_value_label.Text = $"{rumble_duration_milliseconds} ms";
        };
        threshold_trackbar.Scroll += (_, __) =>
        {
            threshold_dbfs = threshold_trackbar.Value;
            threshold_value_label.Text = $"{threshold_trackbar.Value} dBFS";
        };
        attack_time_trackbar.Scroll += (_, __) =>
        {
            attack_time_milliseconds = attack_time_trackbar.Value;
            attack_time_value_label.Text = $"{attack_time_trackbar.Value} ms";
        };
        release_time_trackbar.Scroll += (_, __) =>
        {
            release_time_milliseconds = release_time_trackbar.Value;
            release_time_value_label.Text = $"{release_time_trackbar.Value} ms";
        };
        dwell_time_trackbar.Scroll += (_, __) =>
        {
            dwell_time_milliseconds = dwell_time_trackbar.Value;
            dwell_time_value_label.Text = $"{dwell_time_trackbar.Value} ms";
        };
        cooldown_time_trackbar.Scroll += (_, __) =>
        {
            cooldown_time_milliseconds = cooldown_time_trackbar.Value;
            cooldown_time_value_label.Text = $"{cooldown_time_trackbar.Value} ms";
        };

        // ui timer
        ui_update_timer = new System.Windows.Forms.Timer();
        ui_update_timer.Interval = 50;
        ui_update_timer.Tick += (_, __) => update_ui_from_meter();
        ui_update_timer.Start();

        // nice background tones
        BackColor = Color.FromArgb(245, 247, 250);
        foreach (Control c in Controls) if (c is Panel p) p.BackColor = Color.White;
    }

    // -------------------- UI BUILD HELPERS --------------------
    private GroupBox make_group_box(
        string title,
        out TrackBar trackbar,
        out Label value_label,
        int min,
        int max,
        int initial,
        int tick_frequency,
        string unit_suffix)
    {
        var group = new GroupBox
        {
            Text = title,
            Dock = DockStyle.Fill,                 // FILL the table cell (not Top)
            Padding = new Padding(12),
            MinimumSize = new Size(0, 110)         // safety height
        };

        // inner panel that fills the group
        var inner = new Panel { Dock = DockStyle.Fill };

        trackbar = make_trackbar(0, 20, 480, min, max, initial, tick_frequency);
        trackbar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

        value_label = new Label
        {
            AutoSize = true,
            Left = trackbar.Right + 10,
            Top = 20,
            Text = $"{initial}{unit_suffix}",
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };

        inner.Controls.Add(trackbar);
        inner.Controls.Add(value_label);
        group.Controls.Add(inner);
        return group;
    }


    private TrackBar make_trackbar(int left, int top, int width, int min, int max, int value, int tick_frequency)
    {
        var tb = new TrackBar
        {
            Left = left,
            Top = top,
            Width = width,
            Minimum = min,
            Maximum = max,
            Value = Math.Min(Math.Max(value, min), max),
            TickStyle = TickStyle.BottomRight,
            LargeChange = Math.Max(1, (max - min) / 10),
            SmallChange = 1
        };
        tb.TickFrequency = Math.Max(1, tick_frequency);
        return tb;
    }


    // -------------------- LIFECYCLE --------------------
    private void on_form_load(object sender, EventArgs e)
    {
        audio_device_enumerator = new MMDeviceEnumerator();
        populate_input_devices();
        update_controller_status_label();
    }

    private void on_form_closing(object sender, FormClosingEventArgs e)
    {
        stop_capture_if_active();
        try { set_rumble(0, 0); } catch { }
        audio_device_enumerator?.Dispose();
    }

    // -------------------- DEVICE SELECTION --------------------
    private void populate_input_devices()
    {
        input_device_combo_box.Items.Clear();
        var devices = audio_device_enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active);
        foreach (var device in devices)
        {
            input_device_combo_box.Items.Add(new recording_device_item(device));
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

    // -------------------- START / STOP --------------------
    private void on_start_clicked(object sender, EventArgs e)
    {
        if (wasapi_capture != null) return;

        var selected = input_device_combo_box.SelectedItem as recording_device_item;
        if (selected == null) return;

        if (!xinput_controller.IsConnected)
        {
            update_controller_status_label();
            return;
        }

        wasapi_capture = new WasapiCapture(selected.mm_device, false, capture_buffer_milliseconds);
        wasapi_capture.ShareMode = AudioClientShareMode.Shared;
        wasapi_capture.DataAvailable += on_audio_data_available;
        wasapi_capture.StartRecording();

        start_button.Enabled = false;
        stop_button.Enabled = true;
        input_device_combo_box.Enabled = false;

        smoothed_level_rms_linear = 0.0f;
        consecutive_over_threshold_time_milliseconds = 0;
        last_buffer_time_utc = DateTime.UtcNow;
    }

    private void on_stop_clicked(object sender, EventArgs e)
    {
        stop_capture_if_active();
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

    // -------------------- TEST RUMBLE --------------------
    private void on_test_rumble_clicked(object sender, EventArgs e)
    {
        trigger_rumble_once();
    }

    // -------------------- AUDIO PROCESSING --------------------
    private void on_audio_data_available(object sender, WaveInEventArgs e)
    {
        var format = wasapi_capture.WaveFormat;

        float buffer_rms_linear;

        if (format.Encoding == WaveFormatEncoding.IeeeFloat && format.BitsPerSample == 32)
        {
            int sample_count = e.BytesRecorded / 4;
            double sum_of_squares = 0.0;
            for (int i = 0; i < sample_count; i++)
            {
                float sample = BitConverter.ToSingle(e.Buffer, i * 4);
                sum_of_squares += sample * sample;
            }
            float mean_square = (float)(sum_of_squares / Math.Max(1, sample_count));
            buffer_rms_linear = (float)Math.Sqrt(mean_square);
        }
        else if (format.Encoding == WaveFormatEncoding.Pcm && format.BitsPerSample == 16)
        {
            int sample_count = e.BytesRecorded / 2;
            double sum_of_squares = 0.0;
            for (int i = 0; i < sample_count; i++)
            {
                short s16 = BitConverter.ToInt16(e.Buffer, i * 2);
                float normalized = s16 / 32768.0f;
                sum_of_squares += normalized * normalized;
            }
            float mean_square = (float)(sum_of_squares / Math.Max(1, sample_count));
            buffer_rms_linear = (float)Math.Sqrt(mean_square);
        }
        else
        {
            buffer_rms_linear = 0.0f;
        }

        DateTime now = DateTime.UtcNow;
        double delta_milliseconds = (now - last_buffer_time_utc).TotalMilliseconds;
        last_buffer_time_utc = now;
        if (delta_milliseconds <= 0) delta_milliseconds = capture_buffer_milliseconds;

        double attack_time_constant_ms = Math.Max(1.0, attack_time_milliseconds);
        double release_time_constant_ms = Math.Max(1.0, release_time_milliseconds);
        float attack_alpha = (float)(1.0 - Math.Exp(-delta_milliseconds / attack_time_constant_ms));
        float release_alpha = (float)(1.0 - Math.Exp(-delta_milliseconds / release_time_constant_ms));

        if (buffer_rms_linear > smoothed_level_rms_linear)
        {
            smoothed_level_rms_linear += attack_alpha * (buffer_rms_linear - smoothed_level_rms_linear);
        }
        else
        {
            smoothed_level_rms_linear += release_alpha * (buffer_rms_linear - smoothed_level_rms_linear);
        }

        current_level_rms_linear = smoothed_level_rms_linear;
        current_level_dbfs = (current_level_rms_linear > 0f)
            ? 20.0f * (float)Math.Log10(current_level_rms_linear)
            : -120.0f;

        if (current_level_dbfs >= threshold_dbfs)
        {
            consecutive_over_threshold_time_milliseconds += (int)delta_milliseconds;
        }
        else
        {
            consecutive_over_threshold_time_milliseconds = 0;
        }

        bool cooldown_active = DateTime.UtcNow < last_rumble_end_time_utc.AddMilliseconds(cooldown_time_milliseconds);
        if (!cooldown_active && consecutive_over_threshold_time_milliseconds >= dwell_time_milliseconds)
        {
            trigger_rumble_once();
            consecutive_over_threshold_time_milliseconds = 0;
        }
    }

    // -------------------- UI METER --------------------
    private void update_ui_from_meter()
    {
        float shown_dbfs = float.IsFinite(current_level_dbfs) ? current_level_dbfs : -120.0f;
        current_level_label.Text = $"Current level: {shown_dbfs:0} dBFS";

        int bar_value = (int)Math.Max(0, Math.Min(100, (shown_dbfs + 120f) / 120f * 100f));
        current_level_progress_bar.Value = bar_value;
    }

    // -------------------- RUMBLE --------------------
    private void trigger_rumble_once()
    {
        if (!xinput_controller.IsConnected) return;

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

    // -------------------- HELPER TYPES --------------------
    private class recording_device_item
    {
        public MMDevice mm_device;

        public recording_device_item(MMDevice device)
        {
            mm_device = device;
        }

        public override string ToString()
        {
            return mm_device.FriendlyName;
        }
    }
}
