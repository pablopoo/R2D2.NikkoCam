using System.Drawing;
using System.Windows.Forms;

namespace R2D2.NikkoCam;

// Main operator UI. This class only orchestrates the user-facing flow:
// discover receiver, connect/disconnect video, show preview, and map button
// presses to RC audio tones and channel changes.
internal sealed class MainForm : Form
{
    private sealed record DeviceChoice(string Path, string Label)
    {
        public override string ToString() => Label;
    }

    private sealed class CardPanel : Panel
    {
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using var pen = new Pen(Color.FromArgb(214, 206, 194));
            e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        }
    }

    private sealed record RobotControlSpec(string Key, string Label, string Description, char? DtmfDigit = null);

    private static readonly Color WindowColor = Color.FromArgb(239, 234, 225);
    private static readonly Color CardColor = Color.FromArgb(248, 244, 237);
    private static readonly Color CardBorderColor = Color.FromArgb(185, 193, 201);
    private static readonly Color ButtonColor = Color.FromArgb(229, 234, 240);
    private static readonly Color ButtonBorderColor = Color.FromArgb(185, 193, 201);
    private static readonly Color AccentColor = Color.FromArgb(57, 94, 126);
    private static readonly Color AccentBorderColor = Color.FromArgb(41, 71, 95);
    private static readonly Color AccentTextColor = Color.FromArgb(250, 251, 252);
    private static readonly Color ActiveChannelColor = Color.FromArgb(96, 139, 102);
    private static readonly Color WarmAccentColor = Color.FromArgb(209, 176, 109);
    private static readonly Color WarmAccentBorderColor = Color.FromArgb(161, 131, 74);
    private static readonly Color NoteColor = Color.FromArgb(89, 82, 73);
    private const int SidebarInnerWidth = 318;

    private static readonly RobotControlSpec[] DrivePadControls =
    [
        new("forward-left-turn", "↶", "Left turn.", '1'),
        new("forward", "↑", "Move forward.", '2'),
        new("forward-right-turn", "↷", "Right turn.", '3'),
        new("pivot-left", "⟲", "Counterclockwise pivot turn.", '4'),
        new("camera-reactivation", "Cam", "Camera reactivation. Hold until the screen is restored.", '5'),
        new("pivot-right", "⟳", "Clockwise pivot turn.", '6'),
        new("reverse-left-turn", "↶", "Left turn.", '9'),
        new("reverse", "↓", "Reverse.", '8'),
        new("reverse-right-turn", "↷", "Right turn.", '7'),
    ];

    private static readonly RobotControlSpec[] CameraControls =
    [
        new("camera-up", "Camera Up", "Tilt the camera upward.", '#'),
        new("camera-down", "Camera Down", "Tilt the camera downward.", '*'),
    ];

    private readonly NikkoCameraController _controller = new();
    private readonly RobotRcAudioController _robotRc = new();
    private readonly CancellationTokenSource _lifetimeCts = new();
    private readonly System.Windows.Forms.Timer _rcStatusTimer = new() { Interval = 1000 };
    private readonly ToolTip _toolTip = new()
    {
        AutoPopDelay = 12000,
        InitialDelay = 400,
        ReshowDelay = 150,
    };

    private readonly ComboBox _devicePathCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox _channelCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly Button _refreshDevicesButton = new() { Text = "Refresh" };
    private readonly Button _cameraToggleButton = new() { Text = "Connect" };
    private readonly Button _rcToggleButton = new() { Text = "Connect RC" };

    private readonly PictureBox _previewBox = new()
    {
        Dock = DockStyle.Fill,
        BackColor = Color.Black,
        SizeMode = PictureBoxSizeMode.Zoom,
    };

    private readonly TextBox _logText = new()
    {
        Dock = DockStyle.Fill,
        Multiline = true,
        ReadOnly = true,
        ScrollBars = ScrollBars.Vertical,
        Font = new Font("Consolas", 9f),
        BackColor = Color.FromArgb(252, 248, 240),
    };

    private bool _cameraRunning;
    private bool _busy;
    private int _selectedChannel = 1;
    private bool _lastKnownRcConnected;

    internal MainForm()
    {
        Text = "Nikko R2-D2 Webcam";
        Width = 1440;
        Height = 920;
        MinimumSize = new Size(1180, 780);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = WindowColor;
        Font = new Font("Segoe UI", 9.5f);

        BuildLayout();
        InitializeChannelCombo();
        WireEvents();
        UpdateCameraButtonState();
        UpdateRcButtonState();
        _rcStatusTimer.Start();
    }

    protected override async void OnShown(EventArgs e)
    {
        base.OnShown(e);
        await RefreshDevicesAsync();
    }

    protected override async void OnFormClosing(FormClosingEventArgs e)
    {
        _lifetimeCts.Cancel();
        try
        {
            await _controller.StopAsync();
        }
        catch
        {
        }
        finally
        {
            _rcStatusTimer.Stop();
            _rcStatusTimer.Dispose();
            _robotRc.Dispose();
            _controller.Dispose();
            _previewBox.Image?.Dispose();
            base.OnFormClosing(e);
        }
    }

    private void BuildLayout()
    {
        // Left column: camera and robot controls.
        // Right column: live preview and activity log.
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Padding = new Padding(18),
            BackColor = BackColor,
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 370));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        var left = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 6,
            Padding = new Padding(0),
            Margin = new Padding(0, 0, 16, 0),
        };
        left.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        left.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        left.Controls.Add(BuildCameraCard(), 0, 0);
        left.Controls.Add(BuildRobotControlsCard(), 0, 1);

        var right = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = BackColor,
            Margin = new Padding(0),
        };
        right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        right.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        right.RowStyles.Add(new RowStyle(SizeType.Absolute, 220));

        var previewPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            BackColor = Color.FromArgb(23, 27, 35),
            Margin = new Padding(0),
        };
        previewPanel.Controls.Add(_previewBox);

        var logCardContent = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0),
        };
        logCardContent.Controls.Add(_logText);

        right.Controls.Add(CreateMainCard("Live View", previewPanel), 0, 0);
        right.Controls.Add(CreateMainCard("Activity", logCardContent), 0, 1);

        root.Controls.Add(left, 0, 0);
        root.Controls.Add(right, 1, 0);
        Controls.Add(root);
    }

    private Control BuildReceiverSection()
    {
        _devicePathCombo.Height = 38;
        _devicePathCombo.DropDownWidth = 360;
        _devicePathCombo.BackColor = Color.White;
        _devicePathCombo.Width = 208;

        StyleButton(_refreshDevicesButton, ButtonColor, ButtonBorderColor, Color.FromArgb(48, 55, 63), bold: true);
        _refreshDevicesButton.Size = new Size(92, 38);

        var row = new Panel
        {
            Width = SidebarInnerWidth,
            Height = 38,
            Margin = new Padding(0),
            Padding = new Padding(0),
        };

        _devicePathCombo.Location = new Point(0, 0);
        _refreshDevicesButton.Location = new Point(SidebarInnerWidth - _refreshDevicesButton.Width, 0);

        row.Controls.Add(_devicePathCombo);
        row.Controls.Add(_refreshDevicesButton);
        return row;
    }

    private Control BuildCameraCard()
    {
        StyleButton(_cameraToggleButton, AccentColor, AccentBorderColor, AccentTextColor, bold: true);
        _cameraToggleButton.Height = 44;
        _cameraToggleButton.Width = 180;
        _cameraToggleButton.Location = new Point((SidebarInnerWidth - _cameraToggleButton.Width) / 2, 64);

        var channelLabel = CreateSectionLabel("Channel");
        channelLabel.Location = new Point(0, 126);

        StyleComboBox(_channelCombo);
        _channelCombo.Width = 98;
        _channelCombo.Location = new Point((SidebarInnerWidth - _channelCombo.Width) / 2, 150);
        _channelCombo.Anchor = AnchorStyles.Top;

        var panel = new Panel
        {
            Width = SidebarInnerWidth,
            Height = 190,
            Margin = new Padding(0),
        };
        var receiverSection = BuildReceiverSection();
        receiverSection.Location = new Point(0, 0);
        panel.Controls.Add(receiverSection);
        panel.Controls.Add(CreateSectionLabel("Stream", new Point(0, 42)));
        panel.Controls.Add(_cameraToggleButton);
        panel.Controls.Add(channelLabel);
        panel.Controls.Add(_channelCombo);
        UpdateChannelButtonState();

        return CreateSidebarCard("Camera", panel);
    }

    private Control BuildDriveSection()
    {
        var grid = new TableLayoutPanel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 3,
            RowCount = 3,
            Margin = new Padding(0),
            Padding = new Padding(0),
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.34f));

        for (var index = 0; index < DrivePadControls.Length; index++)
        {
            var spec = DrivePadControls[index];
            var button = new Button
            {
                Text = spec.Label,
                Tag = spec,
                Dock = DockStyle.Fill,
                Height = 72,
                Margin = new Padding(0, 0, 8, 8),
            };
            StyleDriveButton(button, spec);
            _toolTip.SetToolTip(button, spec.Description);
            WireRobotButton(button, spec);
            grid.Controls.Add(button, index % 3, index / 3);
        }

        return grid;
    }

    private Control BuildRobotControlsCard()
    {
        StyleButton(_rcToggleButton, AccentColor, AccentBorderColor, AccentTextColor, bold: true);
        _rcToggleButton.Size = new Size(180, 40);
        _rcToggleButton.Location = new Point((SidebarInnerWidth - _rcToggleButton.Width) / 2, 24);

        var tiltSection = BuildRobotCard("Camera Tilt", CameraControls);
        var tiltSize = tiltSection.GetPreferredSize(new Size(SidebarInnerWidth, 0));
        tiltSection.Size = tiltSize;
        tiltSection.Location = new Point((SidebarInnerWidth - tiltSection.Width) / 2, _rcToggleButton.Bottom + 18);

        var driveLabel = CreateSectionLabel("Drive", new Point(0, tiltSection.Bottom + 10));
        var driveSection = BuildDriveSection();
        var driveSize = driveSection.GetPreferredSize(new Size(SidebarInnerWidth, 0));
        driveSection.Size = driveSize;
        driveSection.Location = new Point((SidebarInnerWidth - driveSection.Width) / 2, driveLabel.Bottom + 6);

        var note = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(320, 0),
            Margin = new Padding(0),
            ForeColor = NoteColor,
            Font = new Font("Segoe UI", 8.75f),
            Text = "Cam reactivates the camera when the robot is in Standby mode on the physical switch. Hold it until the video comes back.",
        };
        note.Size = note.GetPreferredSize(new Size(320, 0));
        note.Location = new Point(0, driveSection.Bottom + 12);

        var panel = new Panel
        {
            Width = SidebarInnerWidth,
            Height = note.Bottom + 6,
            Margin = new Padding(0),
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
        };
        panel.Controls.Add(_rcToggleButton);
        panel.Controls.Add(tiltSection);
        panel.Controls.Add(driveLabel);
        panel.Controls.Add(driveSection);
        panel.Controls.Add(note);

        return CreateSidebarCard("Robot Controls", panel);
    }

    // Wire the UI to the fixed camera/RC actions. Most controls are momentary:
    // press starts the underlying action, release stops it.
    private void WireEvents()
    {
        _refreshDevicesButton.Click += async (_, _) => await RefreshDevicesAsync();
        _cameraToggleButton.Click += async (_, _) => await ToggleCameraAsync();
        _rcToggleButton.Click += (_, _) => ToggleRcConnection();
        _rcStatusTimer.Tick += (_, _) => RefreshRcConnectionUi();
        _channelCombo.SelectionChangeCommitted += async (_, _) =>
        {
            var channelNumber = _channelCombo.SelectedIndex + 1;
            if (channelNumber >= 1)
            {
                await SelectChannelAsync(channelNumber);
            }
        };
    }

    // One toggle button drives the whole connect/disconnect flow.
    private async Task ToggleCameraAsync()
    {
        if (_cameraRunning)
        {
            await StopCameraAsync();
            return;
        }

        await StartCameraAsync();
    }

    // Repopulate the receiver list from the current WinUSB device discovery results.
    private async Task RefreshDevicesAsync()
    {
        SetBusy(true);
        try
        {
            var selectedPath = (_devicePathCombo.SelectedItem as DeviceChoice)?.Path;
            var devices = await Task.Run(_controller.DiscoverDevicePaths);
            _devicePathCombo.BeginUpdate();
            _devicePathCombo.Items.Clear();

            for (var index = 0; index < devices.Count; index++)
            {
                var label = devices.Count == 1
                    ? "Nikko USB Receiver"
                    : $"Nikko USB Receiver {index + 1}";
                _devicePathCombo.Items.Add(new DeviceChoice(devices[index], label));
            }

            if (!string.IsNullOrWhiteSpace(selectedPath))
            {
                for (var index = 0; index < _devicePathCombo.Items.Count; index++)
                {
                    if (_devicePathCombo.Items[index] is DeviceChoice choice &&
                        string.Equals(choice.Path, selectedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        _devicePathCombo.SelectedIndex = index;
                        break;
                    }
                }
            }

            if (_devicePathCombo.SelectedIndex < 0 && devices.Count > 0)
            {
                _devicePathCombo.SelectedIndex = 0;
            }

            _devicePathCombo.EndUpdate();

            if (devices.Count == 0)
            {
                AppendLog("No Nikko receiver found.");
            }
        }
        catch (Exception ex)
        {
            _devicePathCombo.EndUpdate();
            AppendLog($"Could not refresh receivers: {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    // Start the fixed video pipeline and immediately reapply the selected RF channel.
    private async Task StartCameraAsync()
    {
        var devicePath = (_devicePathCombo.SelectedItem as DeviceChoice)?.Path?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(devicePath))
        {
            AppendLog("Select a receiver first.");
            return;
        }

        SetBusy(true);
        try
        {
            AppendLog("Connecting camera...");
            await _controller.StartAsync(devicePath, OnPreviewFrame, _lifetimeCts.Token);
            _cameraRunning = true;
            if (_selectedChannel != 1)
            {
                await _controller.SetChannelAsync(_selectedChannel, _lifetimeCts.Token);
            }
            UpdateSelectedChannelUi();
            AppendLog("Camera connected.");
            UpdateCameraButtonState();
        }
        catch (Exception ex)
        {
            _cameraRunning = false;
            AppendLog($"Could not connect camera: {ex.Message}");
            UpdateCameraButtonState();
        }
        finally
        {
            SetBusy(false);
        }
    }

    // Send the recovered channel command pair while preview stays running.
    private async Task SelectChannelAsync(int channelNumber)
    {
        if (!_cameraRunning)
        {
            return;
        }

        SetBusy(true);
        try
        {
            AppendLog($"Changing to CH{channelNumber}...");
            await _controller.SetChannelAsync(channelNumber, _lifetimeCts.Token);
            _selectedChannel = channelNumber;
            AppendLog($"Channel CH{channelNumber} selected.");
            UpdateSelectedChannelUi();
            UpdateChannelButtonState();
        }
        catch (Exception ex)
        {
            AppendLog($"Could not change channel: {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async Task StopCameraAsync()
    {
        SetBusy(true);
        try
        {
            await _controller.StopAsync();
            _cameraRunning = false;
            AppendLog("Camera disconnected.");
            UpdateCameraButtonState();
        }
        catch (Exception ex)
        {
            AppendLog($"Could not disconnect camera: {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    // Frames arrive on the background preview task. Clone them, marshal to the UI
    // thread, and dispose the previous image to avoid leaking GDI resources.
    private void OnPreviewFrame(PreviewFrameResult frame)
    {
        Bitmap? clone = null;
        try
        {
            clone = (Bitmap)frame.Bitmap.Clone();
        }
        finally
        {
            frame.Bitmap.Dispose();
        }

        if (IsDisposed || Disposing)
        {
            clone.Dispose();
            return;
        }

        try
        {
            BeginInvoke(new Action(() =>
            {
                var old = _previewBox.Image;
                _previewBox.Image = clone;
                old?.Dispose();
            }));
        }
        catch
        {
            clone.Dispose();
        }
    }

    protected override void OnDeactivate(EventArgs e)
    {
        _robotRc.StopTone();
        base.OnDeactivate(e);
    }

    // RC buttons are press-and-hold. MouseDown starts the tone, and any release,
    // capture loss, or pointer exit stops it.
    private void WireRobotButton(Button button, RobotControlSpec spec)
    {
        if (spec.DtmfDigit is char digit)
        {
            button.MouseDown += (_, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    StartDriveTone(spec, digit);
                }
            };
            button.MouseUp += (_, _) => StopDriveTone();
            button.MouseCaptureChanged += (_, _) => StopDriveTone();
            button.MouseLeave += (_, _) =>
            {
                if (Control.MouseButtons == MouseButtons.None)
                {
                    StopDriveTone();
                }
            };
            return;
        }
    }

    private void StartDriveTone(RobotControlSpec spec, char digit)
    {
        try
        {
            _robotRc.StartTone(digit);
            UpdateRcButtonState();
        }
        catch (Exception ex)
        {
            UpdateRcButtonState();
            AppendLog($"RC control is unavailable: {ex.Message}");
        }
    }

    private void StopDriveTone()
    {
        if (!_robotRc.IsToneActive)
        {
            return;
        }

        _robotRc.StopTone();
    }

    // Manual RC reconnect is useful because the receiver audio endpoint can be
    // re-enumerated independently from the video interface after unplug/replug.
    private void ToggleRcConnection()
    {
        try
        {
            if (_robotRc.IsConnected)
            {
                _robotRc.Disconnect();
                AppendLog("RC disconnected.");
            }
            else
            {
                var deviceName = _robotRc.Connect();
                AppendLog($"RC connected: {deviceName}.");
            }
        }
        catch (Exception ex)
        {
            AppendLog($"Could not connect RC: {ex.Message}");
        }
        finally
        {
            _lastKnownRcConnected = _robotRc.RefreshConnectionState();
            UpdateRcButtonState();
        }
    }

    private void SetBusy(bool busy)
    {
        _busy = busy;
        _refreshDevicesButton.Enabled = !busy;
        _cameraToggleButton.Enabled = !busy;
        _devicePathCombo.Enabled = !busy;
        UpdateChannelButtonState();
    }

    private void UpdateCameraButtonState()
    {
        _cameraToggleButton.Text = _cameraRunning ? "Disconnect" : "Connect";
        StyleButton(
            _cameraToggleButton,
            _cameraRunning ? Color.FromArgb(133, 89, 86) : AccentColor,
            _cameraRunning ? Color.FromArgb(103, 67, 65) : AccentBorderColor,
            AccentTextColor,
            bold: true);
        UpdateChannelButtonState();
    }

    private void UpdateRcButtonState()
    {
        _rcToggleButton.Text = _lastKnownRcConnected ? "Disconnect RC" : "Connect RC";
        StyleButton(
            _rcToggleButton,
            _lastKnownRcConnected ? Color.FromArgb(133, 89, 86) : AccentColor,
            _lastKnownRcConnected ? Color.FromArgb(103, 67, 65) : AccentBorderColor,
            AccentTextColor,
            bold: true);
    }

    // Poll the RC endpoint connection state so the button reflects unplug/replug
    // without waiting for the next manual button press.
    private void RefreshRcConnectionUi()
    {
        var isConnected = _robotRc.RefreshConnectionState();
        if (_lastKnownRcConnected && !isConnected)
        {
            AppendLog("RC disconnected.");
        }

        _lastKnownRcConnected = isConnected;
        UpdateRcButtonState();
    }

    private void UpdateChannelButtonState()
    {
        _channelCombo.Enabled = _cameraRunning && !_busy;
        UpdateSelectedChannelUi();
    }

    private void AppendLog(string message)
    {
        if (IsDisposed || Disposing)
        {
            return;
        }

        if (InvokeRequired)
        {
            BeginInvoke(new Action<string>(AppendLog), message);
            return;
        }

        _logText.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
        _logText.SelectionStart = _logText.TextLength;
        _logText.ScrollToCaret();
    }

    private static Control CreateSidebarCard(string title, Control content)
    {
        var layout = new TableLayoutPanel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = Color.Transparent,
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 10),
            ForeColor = Color.FromArgb(72, 65, 57),
            Font = new Font("Segoe UI Semibold", 10.5f),
        }, 0, 0);

        content.Dock = DockStyle.Top;
        content.Margin = new Padding(0);
        layout.Controls.Add(content, 0, 1);

        var card = new CardPanel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(14),
            Margin = new Padding(0, 0, 0, 14),
            Width = 346,
            BackColor = CardColor,
        };
        card.Controls.Add(layout);
        return card;
    }

    private Control BuildRobotCard(string title, IReadOnlyList<RobotControlSpec> controls)
    {
        var wrapper = new Panel
        {
            Margin = new Padding(0, 10, 0, 0),
            Width = 220,
            Height = 78,
        };
        var titleLabel = CreateSectionLabel(title, new Point(0, 0));
        wrapper.Controls.Add(titleLabel);

        var grid = new TableLayoutPanel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 2,
            Location = new Point(0, titleLabel.Bottom + 6),
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));

        for (var index = 0; index < controls.Count; index++)
        {
            var spec = controls[index];
            var button = new Button
            {
                Text = spec.Key == "camera-up" ? "Up" : "Down",
                Tag = spec,
                Dock = DockStyle.Fill,
                Height = 40,
                Margin = new Padding(0, 0, 8, 8),
            };
            StyleButton(button, ButtonColor, ButtonBorderColor, Color.FromArgb(48, 55, 63), bold: true);
            _toolTip.SetToolTip(button, spec.Description);
            WireRobotButton(button, spec);
            grid.Controls.Add(button, index % 2, index / 2);
        }

        wrapper.Controls.Add(grid);
        return wrapper;
    }

    private static Control CreateMainCard(string title, Control content)
    {
        var card = CreateCardBase(title, content);
        card.Dock = DockStyle.Fill;
        card.AutoSize = false;
        card.Margin = new Padding(0, 0, 0, 14);
        return card;
    }

    private static CardPanel CreateCardBase(string title, Control content)
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = Color.Transparent,
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        layout.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 10),
            ForeColor = Color.FromArgb(72, 65, 57),
            Font = new Font("Segoe UI Semibold", 10.5f),
        }, 0, 0);

        content.Dock = DockStyle.Fill;
        content.Margin = new Padding(0);
        layout.Controls.Add(content, 0, 1);

        var card = new CardPanel
        {
            Padding = new Padding(14),
            Margin = new Padding(0, 0, 0, 14),
            BackColor = CardColor,
        };
        card.Controls.Add(layout);
        return card;
    }

    private static void StyleButton(Button button, Color backColor, Color borderColor, Color foreColor, bool bold = false)
    {
        button.UseVisualStyleBackColor = false;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.BorderColor = borderColor;
        button.FlatAppearance.MouseOverBackColor = ControlPaint.Light(backColor, 0.06f);
        button.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(backColor, 0.05f);
        button.BackColor = backColor;
        button.ForeColor = foreColor;
        button.Cursor = Cursors.Hand;
        button.Font = new Font("Segoe UI", 9.5f, bold ? FontStyle.Bold : FontStyle.Regular);
    }

    private static void StyleDriveButton(Button button, RobotControlSpec spec)
    {
        var isCamButton = spec.Key == "camera-reactivation";
        StyleButton(
            button,
            isCamButton ? WarmAccentColor : ButtonColor,
            isCamButton ? WarmAccentBorderColor : ButtonBorderColor,
            isCamButton ? Color.FromArgb(57, 44, 22) : Color.FromArgb(48, 55, 63));
        button.Font = isCamButton
            ? new Font("Segoe UI Semibold", 11f)
            : new Font("Segoe UI Symbol", 20f);
    }

    private void InitializeChannelCombo()
    {
        _channelCombo.Items.Clear();
        _channelCombo.Items.AddRange(["CH1", "CH2", "CH3", "CH4"]);
        _channelCombo.SelectedIndex = _selectedChannel - 1;
    }

    private void UpdateSelectedChannelUi()
    {
        if (_channelCombo.Items.Count == 0)
        {
            StyleComboBox(_channelCombo, Color.White, Color.FromArgb(48, 55, 63));
            return;
        }

        var selectedIndex = Math.Clamp(_selectedChannel - 1, 0, 3);
        if (_channelCombo.SelectedIndex != selectedIndex)
        {
            _channelCombo.SelectedIndex = selectedIndex;
        }

        StyleComboBox(
            _channelCombo,
            _cameraRunning ? ActiveChannelColor : Color.White,
            _cameraRunning ? AccentTextColor : Color.FromArgb(48, 55, 63));
    }

    private static void StyleComboBox(ComboBox comboBox)
    {
        StyleComboBox(comboBox, Color.White, Color.FromArgb(48, 55, 63));
    }

    private static void StyleComboBox(ComboBox comboBox, Color backColor, Color foreColor)
    {
        comboBox.FlatStyle = FlatStyle.Flat;
        comboBox.BackColor = backColor;
        comboBox.ForeColor = foreColor;
        comboBox.Font = new Font("Segoe UI Semibold", 9f);
        comboBox.Margin = new Padding(0);
    }

    private static Label CreateSectionLabel(string text) => CreateSectionLabel(text, Point.Empty);

    private static Label CreateSectionLabel(string text, Point location) => new()
    {
        Text = text,
        AutoSize = true,
        Margin = new Padding(0, 0, 0, 8),
        Location = location,
        ForeColor = Color.FromArgb(93, 86, 78),
        Font = new Font("Segoe UI Semibold", 9f),
    };
}
