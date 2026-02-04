using ReaLTaiizor.Forms;
using ReaLTaiizor.Controls;

namespace MapleStoryMacro
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            txtWindowTitle = new HopeTextBox();
            lblStatus = new Label();
            txtLogDisplay = new ListBox();
            monitorTimer = new System.Windows.Forms.Timer(components);
            btnStartRecording = new CyberButton();
            btnStopRecording = new CyberButton();
            btnRefreshWindow = new CyberButton();
            btnLockWindow = new CyberButton();
            btnSaveScript = new CyberButton();
            btnLoadScript = new CyberButton();
            btnClearEvents = new CyberButton();
            btnViewEvents = new CyberButton();
            btnEditEvents = new CyberButton();
            btnStartPlayback = new CyberButton();
            btnStopPlayback = new CyberButton();
            numPlayTimes = new NumericUpDown();
            lblWindowStatus = new Label();
            lblRecordingStatus = new Label();
            lblPlaybackStatus = new Label();
            btnHotkeySettings = new CyberButton();
            btnCustomKeys = new CyberButton();
            btnScheduler = new CyberButton();
            btnStatistics = new CyberButton();
            btnSaveSettings = new CyberButton();
            btnImportSettings = new CyberButton();
            grpRecording = new CyberGroupBox();
            label2 = new Label();
            grpScript = new CyberGroupBox();
            label3 = new Label();
            grpEvents = new CyberGroupBox();
            label4 = new Label();
            grpPlayback = new CyberGroupBox();
            label1 = new Label();
            lblLoopCount = new Label();
            grpAdvanced = new CyberGroupBox();
            label5 = new Label();
            ((System.ComponentModel.ISupportInitialize)numPlayTimes).BeginInit();
            grpRecording.SuspendLayout();
            grpScript.SuspendLayout();
            grpEvents.SuspendLayout();
            grpPlayback.SuspendLayout();
            grpAdvanced.SuspendLayout();
            SuspendLayout();
            // 
            // txtWindowTitle
            // 
            txtWindowTitle.BackColor = Color.White;
            txtWindowTitle.BaseColor = Color.FromArgb(44, 55, 66);
            txtWindowTitle.BorderColorA = Color.FromArgb(64, 158, 255);
            txtWindowTitle.BorderColorB = Color.FromArgb(220, 223, 230);
            txtWindowTitle.Font = new Font("Segoe UI", 10F);
            txtWindowTitle.ForeColor = Color.FromArgb(48, 49, 51);
            txtWindowTitle.Hint = "";
            txtWindowTitle.Location = new Point(19, 26);
            txtWindowTitle.MaxLength = 32767;
            txtWindowTitle.Multiline = false;
            txtWindowTitle.Name = "txtWindowTitle";
            txtWindowTitle.PasswordChar = '\0';
            txtWindowTitle.ScrollBars = ScrollBars.None;
            txtWindowTitle.SelectedText = "";
            txtWindowTitle.SelectionLength = 0;
            txtWindowTitle.SelectionStart = 0;
            txtWindowTitle.Size = new Size(200, 34);
            txtWindowTitle.TabIndex = 0;
            txtWindowTitle.TabStop = false;
            txtWindowTitle.Text = "MapleStory";
            txtWindowTitle.UseSystemPasswordChar = false;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Font = new Font("Microsoft JhengHei UI", 9F);
            lblStatus.ForeColor = Color.White;
            lblStatus.Location = new Point(20, 69);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(61, 15);
            lblStatus.TabIndex = 1;
            lblStatus.Text = "狀態: 就緒";
            // 
            // txtLogDisplay
            // 
            txtLogDisplay.BackColor = Color.FromArgb(20, 20, 25);
            txtLogDisplay.BorderStyle = BorderStyle.FixedSingle;
            txtLogDisplay.Font = new Font("Consolas", 9F);
            txtLogDisplay.ForeColor = Color.LimeGreen;
            txtLogDisplay.HorizontalScrollbar = true;
            txtLogDisplay.ItemHeight = 14;
            txtLogDisplay.Location = new Point(21, 95);
            txtLogDisplay.Name = "txtLogDisplay";
            txtLogDisplay.SelectionMode = SelectionMode.None;
            txtLogDisplay.Size = new Size(633, 226);
            txtLogDisplay.TabIndex = 1;
            txtLogDisplay.TabStop = false;
            // 
            // btnStartRecording
            // 
            btnStartRecording.Alpha = 20;
            btnStartRecording.BackColor = Color.Transparent;
            btnStartRecording.Background = true;
            btnStartRecording.Background_WidthPen = 4F;
            btnStartRecording.BackgroundPen = true;
            btnStartRecording.ColorBackground = Color.FromArgb(0, 122, 204);
            btnStartRecording.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnStartRecording.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnStartRecording.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnStartRecording.ColorLighting = Color.FromArgb(29, 200, 238);
            btnStartRecording.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnStartRecording.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnStartRecording.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnStartRecording.Effect_1 = true;
            btnStartRecording.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnStartRecording.Effect_1_Transparency = 25;
            btnStartRecording.Effect_2 = true;
            btnStartRecording.Effect_2_ColorBackground = Color.White;
            btnStartRecording.Effect_2_Transparency = 20;
            btnStartRecording.Font = new Font("Microsoft JhengHei UI", 9F);
            btnStartRecording.ForeColor = Color.FromArgb(245, 245, 245);
            btnStartRecording.Lighting = false;
            btnStartRecording.LinearGradient_Background = false;
            btnStartRecording.LinearGradientPen = false;
            btnStartRecording.Location = new Point(10, 35);
            btnStartRecording.Name = "btnStartRecording";
            btnStartRecording.PenWidth = 15;
            btnStartRecording.Rounding = true;
            btnStartRecording.RoundingInt = 70;
            btnStartRecording.Size = new Size(75, 35);
            btnStartRecording.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnStartRecording.TabIndex = 4;
            btnStartRecording.Tag = "Cyber";
            btnStartRecording.TextButton = "▶ 開始";
            btnStartRecording.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnStartRecording.Timer_Effect_1 = 5;
            btnStartRecording.Timer_RGB = 300;
            btnStartRecording.Click += btnStartRecording_Click_1;
            // 
            // btnStopRecording
            // 
            btnStopRecording.Alpha = 20;
            btnStopRecording.BackColor = Color.Transparent;
            btnStopRecording.Background = true;
            btnStopRecording.Background_WidthPen = 4F;
            btnStopRecording.BackgroundPen = true;
            btnStopRecording.ColorBackground = Color.FromArgb(200, 80, 80);
            btnStopRecording.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnStopRecording.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnStopRecording.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnStopRecording.ColorLighting = Color.FromArgb(29, 200, 238);
            btnStopRecording.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnStopRecording.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnStopRecording.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnStopRecording.Effect_1 = true;
            btnStopRecording.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnStopRecording.Effect_1_Transparency = 25;
            btnStopRecording.Effect_2 = true;
            btnStopRecording.Effect_2_ColorBackground = Color.White;
            btnStopRecording.Effect_2_Transparency = 20;
            btnStopRecording.Enabled = false;
            btnStopRecording.Font = new Font("Microsoft JhengHei UI", 9F);
            btnStopRecording.ForeColor = Color.FromArgb(245, 245, 245);
            btnStopRecording.Lighting = false;
            btnStopRecording.LinearGradient_Background = false;
            btnStopRecording.LinearGradientPen = false;
            btnStopRecording.Location = new Point(90, 35);
            btnStopRecording.Name = "btnStopRecording";
            btnStopRecording.PenWidth = 15;
            btnStopRecording.Rounding = true;
            btnStopRecording.RoundingInt = 70;
            btnStopRecording.Size = new Size(75, 35);
            btnStopRecording.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnStopRecording.TabIndex = 5;
            btnStopRecording.Tag = "Cyber";
            btnStopRecording.TextButton = "■ 停止";
            btnStopRecording.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnStopRecording.Timer_Effect_1 = 5;
            btnStopRecording.Timer_RGB = 300;
            // 
            // btnRefreshWindow
            // 
            btnRefreshWindow.Alpha = 20;
            btnRefreshWindow.BackColor = Color.Transparent;
            btnRefreshWindow.Background = true;
            btnRefreshWindow.Background_WidthPen = 4F;
            btnRefreshWindow.BackgroundPen = true;
            btnRefreshWindow.ColorBackground = Color.FromArgb(37, 52, 68);
            btnRefreshWindow.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnRefreshWindow.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnRefreshWindow.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnRefreshWindow.ColorLighting = Color.FromArgb(29, 200, 238);
            btnRefreshWindow.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnRefreshWindow.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnRefreshWindow.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnRefreshWindow.Effect_1 = true;
            btnRefreshWindow.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnRefreshWindow.Effect_1_Transparency = 25;
            btnRefreshWindow.Effect_2 = true;
            btnRefreshWindow.Effect_2_ColorBackground = Color.White;
            btnRefreshWindow.Effect_2_Transparency = 20;
            btnRefreshWindow.Font = new Font("Microsoft JhengHei UI", 9F);
            btnRefreshWindow.ForeColor = Color.FromArgb(245, 245, 245);
            btnRefreshWindow.Lighting = false;
            btnRefreshWindow.LinearGradient_Background = false;
            btnRefreshWindow.LinearGradientPen = false;
            btnRefreshWindow.Location = new Point(231, 27);
            btnRefreshWindow.Name = "btnRefreshWindow";
            btnRefreshWindow.PenWidth = 15;
            btnRefreshWindow.Rounding = true;
            btnRefreshWindow.RoundingInt = 70;
            btnRefreshWindow.Size = new Size(90, 36);
            btnRefreshWindow.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnRefreshWindow.TabIndex = 2;
            btnRefreshWindow.Tag = "Cyber";
            btnRefreshWindow.TextButton = "重新檢測";
            btnRefreshWindow.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnRefreshWindow.Timer_Effect_1 = 5;
            btnRefreshWindow.Timer_RGB = 300;
            // 
            // btnLockWindow
            // 
            btnLockWindow.Alpha = 20;
            btnLockWindow.BackColor = Color.Transparent;
            btnLockWindow.Background = true;
            btnLockWindow.Background_WidthPen = 4F;
            btnLockWindow.BackgroundPen = true;
            btnLockWindow.ColorBackground = Color.FromArgb(37, 52, 68);
            btnLockWindow.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnLockWindow.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnLockWindow.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnLockWindow.ColorLighting = Color.FromArgb(29, 200, 238);
            btnLockWindow.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnLockWindow.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnLockWindow.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnLockWindow.Effect_1 = true;
            btnLockWindow.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnLockWindow.Effect_1_Transparency = 25;
            btnLockWindow.Effect_2 = true;
            btnLockWindow.Effect_2_ColorBackground = Color.White;
            btnLockWindow.Effect_2_Transparency = 20;
            btnLockWindow.Font = new Font("Microsoft JhengHei UI", 9F);
            btnLockWindow.ForeColor = Color.FromArgb(245, 245, 245);
            btnLockWindow.Lighting = false;
            btnLockWindow.LinearGradient_Background = false;
            btnLockWindow.LinearGradientPen = false;
            btnLockWindow.Location = new Point(339, 27);
            btnLockWindow.Name = "btnLockWindow";
            btnLockWindow.PenWidth = 15;
            btnLockWindow.Rounding = true;
            btnLockWindow.RoundingInt = 70;
            btnLockWindow.Size = new Size(90, 36);
            btnLockWindow.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnLockWindow.TabIndex = 3;
            btnLockWindow.Tag = "Cyber";
            btnLockWindow.TextButton = "手動鎖定";
            btnLockWindow.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnLockWindow.Timer_Effect_1 = 5;
            btnLockWindow.Timer_RGB = 300;
            // 
            // btnSaveScript
            // 
            btnSaveScript.Alpha = 20;
            btnSaveScript.BackColor = Color.Transparent;
            btnSaveScript.Background = true;
            btnSaveScript.Background_WidthPen = 4F;
            btnSaveScript.BackgroundPen = true;
            btnSaveScript.ColorBackground = Color.FromArgb(60, 120, 80);
            btnSaveScript.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnSaveScript.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnSaveScript.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnSaveScript.ColorLighting = Color.FromArgb(29, 200, 238);
            btnSaveScript.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnSaveScript.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnSaveScript.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnSaveScript.Effect_1 = true;
            btnSaveScript.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnSaveScript.Effect_1_Transparency = 25;
            btnSaveScript.Effect_2 = true;
            btnSaveScript.Effect_2_ColorBackground = Color.White;
            btnSaveScript.Effect_2_Transparency = 20;
            btnSaveScript.Font = new Font("Microsoft JhengHei UI", 9F);
            btnSaveScript.ForeColor = Color.FromArgb(245, 245, 245);
            btnSaveScript.Lighting = false;
            btnSaveScript.LinearGradient_Background = false;
            btnSaveScript.LinearGradientPen = false;
            btnSaveScript.Location = new Point(10, 35);
            btnSaveScript.Name = "btnSaveScript";
            btnSaveScript.PenWidth = 15;
            btnSaveScript.Rounding = true;
            btnSaveScript.RoundingInt = 70;
            btnSaveScript.Size = new Size(68, 35);
            btnSaveScript.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnSaveScript.TabIndex = 6;
            btnSaveScript.Tag = "Cyber";
            btnSaveScript.TextButton = "保存";
            btnSaveScript.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnSaveScript.Timer_Effect_1 = 5;
            btnSaveScript.Timer_RGB = 300;
            // 
            // btnLoadScript
            // 
            btnLoadScript.Alpha = 20;
            btnLoadScript.BackColor = Color.Transparent;
            btnLoadScript.Background = true;
            btnLoadScript.Background_WidthPen = 4F;
            btnLoadScript.BackgroundPen = true;
            btnLoadScript.ColorBackground = Color.FromArgb(80, 100, 140);
            btnLoadScript.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnLoadScript.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnLoadScript.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnLoadScript.ColorLighting = Color.FromArgb(29, 200, 238);
            btnLoadScript.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnLoadScript.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnLoadScript.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnLoadScript.Effect_1 = true;
            btnLoadScript.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnLoadScript.Effect_1_Transparency = 25;
            btnLoadScript.Effect_2 = true;
            btnLoadScript.Effect_2_ColorBackground = Color.White;
            btnLoadScript.Effect_2_Transparency = 20;
            btnLoadScript.Font = new Font("Microsoft JhengHei UI", 9F);
            btnLoadScript.ForeColor = Color.FromArgb(245, 245, 245);
            btnLoadScript.Lighting = false;
            btnLoadScript.LinearGradient_Background = false;
            btnLoadScript.LinearGradientPen = false;
            btnLoadScript.Location = new Point(83, 35);
            btnLoadScript.Name = "btnLoadScript";
            btnLoadScript.PenWidth = 15;
            btnLoadScript.Rounding = true;
            btnLoadScript.RoundingInt = 70;
            btnLoadScript.Size = new Size(68, 35);
            btnLoadScript.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnLoadScript.TabIndex = 7;
            btnLoadScript.Tag = "Cyber";
            btnLoadScript.TextButton = "載入";
            btnLoadScript.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnLoadScript.Timer_Effect_1 = 5;
            btnLoadScript.Timer_RGB = 300;
            // 
            // btnClearEvents
            // 
            btnClearEvents.Alpha = 20;
            btnClearEvents.BackColor = Color.Transparent;
            btnClearEvents.Background = true;
            btnClearEvents.Background_WidthPen = 4F;
            btnClearEvents.BackgroundPen = true;
            btnClearEvents.ColorBackground = Color.FromArgb(150, 60, 60);
            btnClearEvents.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnClearEvents.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnClearEvents.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnClearEvents.ColorLighting = Color.FromArgb(29, 200, 238);
            btnClearEvents.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnClearEvents.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnClearEvents.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnClearEvents.Effect_1 = true;
            btnClearEvents.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnClearEvents.Effect_1_Transparency = 25;
            btnClearEvents.Effect_2 = true;
            btnClearEvents.Effect_2_ColorBackground = Color.White;
            btnClearEvents.Effect_2_Transparency = 20;
            btnClearEvents.Font = new Font("Microsoft JhengHei UI", 9F);
            btnClearEvents.ForeColor = Color.FromArgb(245, 245, 245);
            btnClearEvents.Lighting = false;
            btnClearEvents.LinearGradient_Background = false;
            btnClearEvents.LinearGradientPen = false;
            btnClearEvents.Location = new Point(101, 35);
            btnClearEvents.Name = "btnClearEvents";
            btnClearEvents.PenWidth = 15;
            btnClearEvents.Rounding = true;
            btnClearEvents.RoundingInt = 70;
            btnClearEvents.Size = new Size(60, 35);
            btnClearEvents.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnClearEvents.TabIndex = 8;
            btnClearEvents.Tag = "Cyber";
            btnClearEvents.TextButton = "清除";
            btnClearEvents.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnClearEvents.Timer_Effect_1 = 5;
            btnClearEvents.Timer_RGB = 300;
            // 
            // btnViewEvents
            // 
            btnViewEvents.Alpha = 20;
            btnViewEvents.BackColor = Color.Transparent;
            btnViewEvents.Background = true;
            btnViewEvents.Background_WidthPen = 4F;
            btnViewEvents.BackgroundPen = true;
            btnViewEvents.ColorBackground = Color.FromArgb(70, 90, 110);
            btnViewEvents.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnViewEvents.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnViewEvents.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnViewEvents.ColorLighting = Color.FromArgb(29, 200, 238);
            btnViewEvents.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnViewEvents.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnViewEvents.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnViewEvents.Effect_1 = true;
            btnViewEvents.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnViewEvents.Effect_1_Transparency = 25;
            btnViewEvents.Effect_2 = true;
            btnViewEvents.Effect_2_ColorBackground = Color.White;
            btnViewEvents.Effect_2_Transparency = 20;
            btnViewEvents.Font = new Font("Microsoft JhengHei UI", 9F);
            btnViewEvents.ForeColor = Color.FromArgb(245, 245, 245);
            btnViewEvents.Lighting = false;
            btnViewEvents.LinearGradient_Background = false;
            btnViewEvents.LinearGradientPen = false;
            btnViewEvents.Location = new Point(40, 35);
            btnViewEvents.Name = "btnViewEvents";
            btnViewEvents.PenWidth = 15;
            btnViewEvents.Rounding = true;
            btnViewEvents.RoundingInt = 70;
            btnViewEvents.Size = new Size(55, 35);
            btnViewEvents.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnViewEvents.TabIndex = 9;
            btnViewEvents.Tag = "Cyber";
            btnViewEvents.TextButton = "查看";
            btnViewEvents.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnViewEvents.Timer_Effect_1 = 5;
            btnViewEvents.Timer_RGB = 300;
            // 
            // btnEditEvents
            // 
            btnEditEvents.Alpha = 20;
            btnEditEvents.BackColor = Color.Transparent;
            btnEditEvents.Background = true;
            btnEditEvents.Background_WidthPen = 4F;
            btnEditEvents.BackgroundPen = true;
            btnEditEvents.ColorBackground = Color.FromArgb(120, 100, 60);
            btnEditEvents.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnEditEvents.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnEditEvents.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnEditEvents.ColorLighting = Color.FromArgb(29, 200, 238);
            btnEditEvents.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnEditEvents.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnEditEvents.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnEditEvents.Effect_1 = true;
            btnEditEvents.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnEditEvents.Effect_1_Transparency = 25;
            btnEditEvents.Effect_2 = true;
            btnEditEvents.Effect_2_ColorBackground = Color.White;
            btnEditEvents.Effect_2_Transparency = 20;
            btnEditEvents.Font = new Font("Microsoft JhengHei UI", 9F);
            btnEditEvents.ForeColor = Color.FromArgb(245, 245, 245);
            btnEditEvents.Lighting = false;
            btnEditEvents.LinearGradient_Background = false;
            btnEditEvents.LinearGradientPen = false;
            btnEditEvents.Location = new Point(156, 35);
            btnEditEvents.Name = "btnEditEvents";
            btnEditEvents.PenWidth = 15;
            btnEditEvents.Rounding = true;
            btnEditEvents.RoundingInt = 70;
            btnEditEvents.Size = new Size(68, 35);
            btnEditEvents.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnEditEvents.TabIndex = 10;
            btnEditEvents.Tag = "Cyber";
            btnEditEvents.TextButton = "編輯";
            btnEditEvents.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnEditEvents.Timer_Effect_1 = 5;
            btnEditEvents.Timer_RGB = 300;
            // 
            // btnStartPlayback
            // 
            btnStartPlayback.Alpha = 20;
            btnStartPlayback.BackColor = Color.Transparent;
            btnStartPlayback.Background = true;
            btnStartPlayback.Background_WidthPen = 4F;
            btnStartPlayback.BackgroundPen = true;
            btnStartPlayback.ColorBackground = Color.FromArgb(0, 150, 80);
            btnStartPlayback.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnStartPlayback.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnStartPlayback.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnStartPlayback.ColorLighting = Color.FromArgb(29, 200, 238);
            btnStartPlayback.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnStartPlayback.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnStartPlayback.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnStartPlayback.Effect_1 = true;
            btnStartPlayback.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnStartPlayback.Effect_1_Transparency = 25;
            btnStartPlayback.Effect_2 = true;
            btnStartPlayback.Effect_2_ColorBackground = Color.White;
            btnStartPlayback.Effect_2_Transparency = 20;
            btnStartPlayback.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
            btnStartPlayback.ForeColor = Color.FromArgb(245, 245, 245);
            btnStartPlayback.Lighting = false;
            btnStartPlayback.LinearGradient_Background = false;
            btnStartPlayback.LinearGradientPen = false;
            btnStartPlayback.Location = new Point(10, 35);
            btnStartPlayback.Name = "btnStartPlayback";
            btnStartPlayback.PenWidth = 15;
            btnStartPlayback.Rounding = true;
            btnStartPlayback.RoundingInt = 70;
            btnStartPlayback.Size = new Size(90, 35);
            btnStartPlayback.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnStartPlayback.TabIndex = 11;
            btnStartPlayback.Tag = "Cyber";
            btnStartPlayback.TextButton = "▶ 播放";
            btnStartPlayback.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnStartPlayback.Timer_Effect_1 = 5;
            btnStartPlayback.Timer_RGB = 300;
            // 
            // btnStopPlayback
            // 
            btnStopPlayback.Alpha = 20;
            btnStopPlayback.BackColor = Color.Transparent;
            btnStopPlayback.Background = true;
            btnStopPlayback.Background_WidthPen = 4F;
            btnStopPlayback.BackgroundPen = true;
            btnStopPlayback.ColorBackground = Color.FromArgb(180, 60, 60);
            btnStopPlayback.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnStopPlayback.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnStopPlayback.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnStopPlayback.ColorLighting = Color.FromArgb(29, 200, 238);
            btnStopPlayback.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnStopPlayback.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnStopPlayback.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnStopPlayback.Effect_1 = true;
            btnStopPlayback.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnStopPlayback.Effect_1_Transparency = 25;
            btnStopPlayback.Effect_2 = true;
            btnStopPlayback.Effect_2_ColorBackground = Color.White;
            btnStopPlayback.Effect_2_Transparency = 20;
            btnStopPlayback.Enabled = false;
            btnStopPlayback.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
            btnStopPlayback.ForeColor = Color.FromArgb(245, 245, 245);
            btnStopPlayback.Lighting = false;
            btnStopPlayback.LinearGradient_Background = false;
            btnStopPlayback.LinearGradientPen = false;
            btnStopPlayback.Location = new Point(110, 35);
            btnStopPlayback.Name = "btnStopPlayback";
            btnStopPlayback.PenWidth = 15;
            btnStopPlayback.Rounding = true;
            btnStopPlayback.RoundingInt = 70;
            btnStopPlayback.Size = new Size(90, 35);
            btnStopPlayback.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnStopPlayback.TabIndex = 12;
            btnStopPlayback.Tag = "Cyber";
            btnStopPlayback.TextButton = "■ 停止";
            btnStopPlayback.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnStopPlayback.Timer_Effect_1 = 5;
            btnStopPlayback.Timer_RGB = 300;
            // 
            // numPlayTimes
            // 
            numPlayTimes.BackColor = Color.FromArgb(60, 60, 65);
            numPlayTimes.Font = new Font("Segoe UI", 10F);
            numPlayTimes.ForeColor = Color.White;
            numPlayTimes.Location = new Point(280, 41);
            numPlayTimes.Maximum = new decimal(new int[] { 9999, 0, 0, 0 });
            numPlayTimes.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            numPlayTimes.Name = "numPlayTimes";
            numPlayTimes.Size = new Size(60, 25);
            numPlayTimes.TabIndex = 13;
            numPlayTimes.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // lblWindowStatus
            // 
            lblWindowStatus.AutoSize = true;
            lblWindowStatus.Font = new Font("Microsoft JhengHei UI", 9F);
            lblWindowStatus.ForeColor = Color.LightGray;
            lblWindowStatus.Location = new Point(21, 519);
            lblWindowStatus.Name = "lblWindowStatus";
            lblWindowStatus.Size = new Size(106, 15);
            lblWindowStatus.TabIndex = 14;
            lblWindowStatus.Text = "視窗狀態: 尋找中...";
            // 
            // lblRecordingStatus
            // 
            lblRecordingStatus.AutoSize = true;
            lblRecordingStatus.Font = new Font("Microsoft JhengHei UI", 9F);
            lblRecordingStatus.ForeColor = Color.LightGray;
            lblRecordingStatus.Location = new Point(21, 539);
            lblRecordingStatus.Name = "lblRecordingStatus";
            lblRecordingStatus.Size = new Size(143, 15);
            lblRecordingStatus.TabIndex = 15;
            lblRecordingStatus.Text = "錄製狀態: 就緒 | 事件數: 0";
            // 
            // lblPlaybackStatus
            // 
            lblPlaybackStatus.AutoSize = true;
            lblPlaybackStatus.Font = new Font("Microsoft JhengHei UI", 9F);
            lblPlaybackStatus.ForeColor = Color.LightGray;
            lblPlaybackStatus.Location = new Point(21, 559);
            lblPlaybackStatus.Name = "lblPlaybackStatus";
            lblPlaybackStatus.Size = new Size(85, 15);
            lblPlaybackStatus.TabIndex = 16;
            lblPlaybackStatus.Text = "播放狀態: 就緒";
            // 
            // btnHotkeySettings
            // 
            btnHotkeySettings.Alpha = 20;
            btnHotkeySettings.BackColor = Color.Transparent;
            btnHotkeySettings.Background = true;
            btnHotkeySettings.Background_WidthPen = 4F;
            btnHotkeySettings.BackgroundPen = true;
            btnHotkeySettings.ColorBackground = Color.FromArgb(100, 80, 120);
            btnHotkeySettings.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnHotkeySettings.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnHotkeySettings.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnHotkeySettings.ColorLighting = Color.FromArgb(29, 200, 238);
            btnHotkeySettings.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnHotkeySettings.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnHotkeySettings.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnHotkeySettings.Effect_1 = true;
            btnHotkeySettings.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnHotkeySettings.Effect_1_Transparency = 25;
            btnHotkeySettings.Effect_2 = true;
            btnHotkeySettings.Effect_2_ColorBackground = Color.White;
            btnHotkeySettings.Effect_2_Transparency = 20;
            btnHotkeySettings.Font = new Font("Microsoft JhengHei UI", 9F);
            btnHotkeySettings.ForeColor = Color.FromArgb(245, 245, 245);
            btnHotkeySettings.Lighting = false;
            btnHotkeySettings.LinearGradient_Background = false;
            btnHotkeySettings.LinearGradientPen = false;
            btnHotkeySettings.Location = new Point(444, 27);
            btnHotkeySettings.Name = "btnHotkeySettings";
            btnHotkeySettings.PenWidth = 15;
            btnHotkeySettings.Rounding = true;
            btnHotkeySettings.RoundingInt = 70;
            btnHotkeySettings.Size = new Size(100, 36);
            btnHotkeySettings.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnHotkeySettings.TabIndex = 17;
            btnHotkeySettings.Tag = "Cyber";
            btnHotkeySettings.TextButton = "⚙ 熱鍵設定";
            btnHotkeySettings.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnHotkeySettings.Timer_Effect_1 = 5;
            btnHotkeySettings.Timer_RGB = 300;
            // 
            // btnCustomKeys
            // 
            btnCustomKeys.Alpha = 20;
            btnCustomKeys.BackColor = Color.Transparent;
            btnCustomKeys.Background = true;
            btnCustomKeys.Background_WidthPen = 4F;
            btnCustomKeys.BackgroundPen = true;
            btnCustomKeys.ColorBackground = Color.FromArgb(80, 120, 80);
            btnCustomKeys.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnCustomKeys.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnCustomKeys.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnCustomKeys.ColorLighting = Color.FromArgb(29, 200, 238);
            btnCustomKeys.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnCustomKeys.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnCustomKeys.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnCustomKeys.Effect_1 = true;
            btnCustomKeys.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnCustomKeys.Effect_1_Transparency = 25;
            btnCustomKeys.Effect_2 = true;
            btnCustomKeys.Effect_2_ColorBackground = Color.White;
            btnCustomKeys.Effect_2_Transparency = 20;
            btnCustomKeys.Font = new Font("Microsoft JhengHei UI", 9F);
            btnCustomKeys.ForeColor = Color.FromArgb(245, 245, 245);
            btnCustomKeys.Lighting = false;
            btnCustomKeys.LinearGradient_Background = false;
            btnCustomKeys.LinearGradientPen = false;
            btnCustomKeys.Location = new Point(12, 35);
            btnCustomKeys.Name = "btnCustomKeys";
            btnCustomKeys.PenWidth = 15;
            btnCustomKeys.Rounding = true;
            btnCustomKeys.RoundingInt = 70;
            btnCustomKeys.Size = new Size(90, 35);
            btnCustomKeys.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnCustomKeys.TabIndex = 24;
            btnCustomKeys.Tag = "Cyber";
            btnCustomKeys.TextButton = "⚡ 自定義";
            btnCustomKeys.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnCustomKeys.Timer_Effect_1 = 5;
            btnCustomKeys.Timer_RGB = 300;
            btnCustomKeys.Click += btnCustomKeys_Click;
            // 
            // btnScheduler
            // 
            btnScheduler.Alpha = 20;
            btnScheduler.BackColor = Color.Transparent;
            btnScheduler.Background = true;
            btnScheduler.Background_WidthPen = 4F;
            btnScheduler.BackgroundPen = true;
            btnScheduler.ColorBackground = Color.FromArgb(120, 100, 60);
            btnScheduler.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnScheduler.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnScheduler.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnScheduler.ColorLighting = Color.FromArgb(29, 200, 238);
            btnScheduler.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnScheduler.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnScheduler.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnScheduler.Effect_1 = true;
            btnScheduler.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnScheduler.Effect_1_Transparency = 25;
            btnScheduler.Effect_2 = true;
            btnScheduler.Effect_2_ColorBackground = Color.White;
            btnScheduler.Effect_2_Transparency = 20;
            btnScheduler.Font = new Font("Microsoft JhengHei UI", 9F);
            btnScheduler.ForeColor = Color.FromArgb(245, 245, 245);
            btnScheduler.Lighting = false;
            btnScheduler.LinearGradient_Background = false;
            btnScheduler.LinearGradientPen = false;
            btnScheduler.Location = new Point(107, 35);
            btnScheduler.Name = "btnScheduler";
            btnScheduler.PenWidth = 15;
            btnScheduler.Rounding = true;
            btnScheduler.RoundingInt = 70;
            btnScheduler.Size = new Size(80, 35);
            btnScheduler.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnScheduler.TabIndex = 25;
            btnScheduler.Tag = "Cyber";
            btnScheduler.TextButton = "⏰ 定時";
            btnScheduler.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnScheduler.Timer_Effect_1 = 5;
            btnScheduler.Timer_RGB = 300;
            // 
            // btnStatistics
            // 
            btnStatistics.Alpha = 20;
            btnStatistics.BackColor = Color.Transparent;
            btnStatistics.Background = true;
            btnStatistics.Background_WidthPen = 4F;
            btnStatistics.BackgroundPen = true;
            btnStatistics.ColorBackground = Color.FromArgb(70, 90, 130);
            btnStatistics.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnStatistics.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnStatistics.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnStatistics.ColorLighting = Color.FromArgb(29, 200, 238);
            btnStatistics.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnStatistics.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnStatistics.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnStatistics.Effect_1 = true;
            btnStatistics.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnStatistics.Effect_1_Transparency = 25;
            btnStatistics.Effect_2 = true;
            btnStatistics.Effect_2_ColorBackground = Color.White;
            btnStatistics.Effect_2_Transparency = 20;
            btnStatistics.Font = new Font("Microsoft JhengHei UI", 9F);
            btnStatistics.ForeColor = Color.FromArgb(245, 245, 245);
            btnStatistics.Lighting = false;
            btnStatistics.LinearGradient_Background = false;
            btnStatistics.LinearGradientPen = false;
            btnStatistics.Location = new Point(192, 35);
            btnStatistics.Name = "btnStatistics";
            btnStatistics.PenWidth = 15;
            btnStatistics.Rounding = true;
            btnStatistics.RoundingInt = 70;
            btnStatistics.Size = new Size(70, 35);
            btnStatistics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnStatistics.TabIndex = 26;
            btnStatistics.Tag = "Cyber";
            btnStatistics.TextButton = "O 統計";
            btnStatistics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnStatistics.Timer_Effect_1 = 5;
            btnStatistics.Timer_RGB = 300;
            // 
            // btnSaveSettings
            // 
            btnSaveSettings.Alpha = 20;
            btnSaveSettings.BackColor = Color.Transparent;
            btnSaveSettings.Background = true;
            btnSaveSettings.Background_WidthPen = 4F;
            btnSaveSettings.BackgroundPen = true;
            btnSaveSettings.ColorBackground = Color.FromArgb(60, 120, 80);
            btnSaveSettings.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnSaveSettings.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnSaveSettings.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnSaveSettings.ColorLighting = Color.FromArgb(29, 200, 238);
            btnSaveSettings.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnSaveSettings.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnSaveSettings.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnSaveSettings.Effect_1 = true;
            btnSaveSettings.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnSaveSettings.Effect_1_Transparency = 25;
            btnSaveSettings.Effect_2 = true;
            btnSaveSettings.Effect_2_ColorBackground = Color.White;
            btnSaveSettings.Effect_2_Transparency = 20;
            btnSaveSettings.Font = new Font("Microsoft JhengHei UI", 9F);
            btnSaveSettings.ForeColor = Color.FromArgb(245, 245, 245);
            btnSaveSettings.Lighting = false;
            btnSaveSettings.LinearGradient_Background = false;
            btnSaveSettings.LinearGradientPen = false;
            btnSaveSettings.Location = new Point(560, 11);
            btnSaveSettings.Name = "btnSaveSettings";
            btnSaveSettings.PenWidth = 15;
            btnSaveSettings.Rounding = true;
            btnSaveSettings.RoundingInt = 70;
            btnSaveSettings.Size = new Size(99, 35);
            btnSaveSettings.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnSaveSettings.TabIndex = 28;
            btnSaveSettings.Tag = "Cyber";
            btnSaveSettings.TextButton = "▼ 導出設定";
            btnSaveSettings.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnSaveSettings.Timer_Effect_1 = 5;
            btnSaveSettings.Timer_RGB = 300;
            btnSaveSettings.Click += btnSaveSettings_Click;
            // 
            // btnImportSettings
            // 
            btnImportSettings.Alpha = 20;
            btnImportSettings.BackColor = Color.Transparent;
            btnImportSettings.Background = true;
            btnImportSettings.Background_WidthPen = 4F;
            btnImportSettings.BackgroundPen = true;
            btnImportSettings.ColorBackground = Color.FromArgb(70, 90, 130);
            btnImportSettings.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            btnImportSettings.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            btnImportSettings.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            btnImportSettings.ColorLighting = Color.FromArgb(29, 200, 238);
            btnImportSettings.ColorPen_1 = Color.FromArgb(37, 52, 68);
            btnImportSettings.ColorPen_2 = Color.FromArgb(41, 63, 86);
            btnImportSettings.CyberButtonStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            btnImportSettings.Effect_1 = true;
            btnImportSettings.Effect_1_ColorBackground = Color.FromArgb(29, 200, 238);
            btnImportSettings.Effect_1_Transparency = 25;
            btnImportSettings.Effect_2 = true;
            btnImportSettings.Effect_2_ColorBackground = Color.White;
            btnImportSettings.Effect_2_Transparency = 20;
            btnImportSettings.Font = new Font("Microsoft JhengHei UI", 9F);
            btnImportSettings.ForeColor = Color.FromArgb(245, 245, 245);
            btnImportSettings.Lighting = false;
            btnImportSettings.LinearGradient_Background = false;
            btnImportSettings.LinearGradientPen = false;
            btnImportSettings.Location = new Point(560, 50);
            btnImportSettings.Name = "btnImportSettings";
            btnImportSettings.PenWidth = 15;
            btnImportSettings.Rounding = true;
            btnImportSettings.RoundingInt = 70;
            btnImportSettings.Size = new Size(99, 35);
            btnImportSettings.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            btnImportSettings.TabIndex = 29;
            btnImportSettings.Tag = "Cyber";
            btnImportSettings.TextButton = "▲ 導入設定";
            btnImportSettings.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            btnImportSettings.Timer_Effect_1 = 5;
            btnImportSettings.Timer_RGB = 300;
            btnImportSettings.Click += btnImportSettings_Click;
            // 
            // grpRecording
            // 
            grpRecording.Alpha = 20;
            grpRecording.BackColor = Color.Transparent;
            grpRecording.Background = true;
            grpRecording.Background_WidthPen = 3F;
            grpRecording.BackgroundPen = true;
            grpRecording.ColorBackground = Color.FromArgb(45, 45, 48);
            grpRecording.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            grpRecording.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            grpRecording.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            grpRecording.ColorLighting = Color.FromArgb(29, 200, 238);
            grpRecording.ColorPen_1 = Color.FromArgb(37, 52, 68);
            grpRecording.ColorPen_2 = Color.FromArgb(41, 63, 86);
            grpRecording.Controls.Add(label2);
            grpRecording.Controls.Add(btnStartRecording);
            grpRecording.Controls.Add(btnStopRecording);
            grpRecording.CyberGroupBoxStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            grpRecording.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
            grpRecording.ForeColor = Color.White;
            grpRecording.Lighting = false;
            grpRecording.LinearGradient_Background = false;
            grpRecording.LinearGradientPen = false;
            grpRecording.Location = new Point(21, 339);
            grpRecording.Name = "grpRecording";
            grpRecording.PenWidth = 15;
            grpRecording.RGB = false;
            grpRecording.Rounding = true;
            grpRecording.RoundingInt = 12;
            grpRecording.Size = new Size(175, 80);
            grpRecording.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            grpRecording.TabIndex = 20;
            grpRecording.TabStop = false;
            grpRecording.Tag = "Cyber";
            grpRecording.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            grpRecording.Timer_RGB = 300;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(60, 11);
            label2.Name = "label2";
            label2.Size = new Size(55, 15);
            label2.TabIndex = 6;
            label2.Text = "錄製控制";
            // 
            // grpScript
            // 
            grpScript.Alpha = 20;
            grpScript.BackColor = Color.Transparent;
            grpScript.Background = true;
            grpScript.Background_WidthPen = 3F;
            grpScript.BackgroundPen = true;
            grpScript.ColorBackground = Color.FromArgb(45, 45, 48);
            grpScript.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            grpScript.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            grpScript.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            grpScript.ColorLighting = Color.FromArgb(29, 200, 238);
            grpScript.ColorPen_1 = Color.FromArgb(37, 52, 68);
            grpScript.ColorPen_2 = Color.FromArgb(41, 63, 86);
            grpScript.Controls.Add(label3);
            grpScript.Controls.Add(btnSaveScript);
            grpScript.Controls.Add(btnLoadScript);
            grpScript.Controls.Add(btnEditEvents);
            grpScript.CyberGroupBoxStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            grpScript.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
            grpScript.ForeColor = Color.White;
            grpScript.Lighting = false;
            grpScript.LinearGradient_Background = false;
            grpScript.LinearGradientPen = false;
            grpScript.Location = new Point(206, 339);
            grpScript.Name = "grpScript";
            grpScript.PenWidth = 15;
            grpScript.RGB = false;
            grpScript.Rounding = true;
            grpScript.RoundingInt = 12;
            grpScript.Size = new Size(235, 80);
            grpScript.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            grpScript.TabIndex = 21;
            grpScript.TabStop = false;
            grpScript.Tag = "Cyber";
            grpScript.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            grpScript.Timer_RGB = 300;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(91, 11);
            label3.Name = "label3";
            label3.Size = new Size(55, 15);
            label3.TabIndex = 11;
            label3.Text = "腳本管理";
            // 
            // grpEvents
            // 
            grpEvents.Alpha = 20;
            grpEvents.BackColor = Color.Transparent;
            grpEvents.Background = true;
            grpEvents.Background_WidthPen = 3F;
            grpEvents.BackgroundPen = true;
            grpEvents.ColorBackground = Color.FromArgb(45, 45, 48);
            grpEvents.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            grpEvents.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            grpEvents.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            grpEvents.ColorLighting = Color.FromArgb(29, 200, 238);
            grpEvents.ColorPen_1 = Color.FromArgb(37, 52, 68);
            grpEvents.ColorPen_2 = Color.FromArgb(41, 63, 86);
            grpEvents.Controls.Add(label4);
            grpEvents.Controls.Add(btnViewEvents);
            grpEvents.Controls.Add(btnClearEvents);
            grpEvents.CyberGroupBoxStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            grpEvents.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
            grpEvents.ForeColor = Color.White;
            grpEvents.Lighting = false;
            grpEvents.LinearGradient_Background = false;
            grpEvents.LinearGradientPen = false;
            grpEvents.Location = new Point(451, 339);
            grpEvents.Name = "grpEvents";
            grpEvents.PenWidth = 15;
            grpEvents.RGB = false;
            grpEvents.Rounding = true;
            grpEvents.RoundingInt = 12;
            grpEvents.Size = new Size(203, 80);
            grpEvents.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            grpEvents.TabIndex = 22;
            grpEvents.TabStop = false;
            grpEvents.Tag = "Cyber";
            grpEvents.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            grpEvents.Timer_RGB = 300;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(72, 11);
            label4.Name = "label4";
            label4.Size = new Size(55, 15);
            label4.TabIndex = 10;
            label4.Text = "事件管理";
            // 
            // grpPlayback
            // 
            grpPlayback.Alpha = 20;
            grpPlayback.BackColor = Color.Transparent;
            grpPlayback.Background = true;
            grpPlayback.Background_WidthPen = 3F;
            grpPlayback.BackgroundPen = true;
            grpPlayback.ColorBackground = Color.FromArgb(45, 45, 48);
            grpPlayback.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            grpPlayback.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            grpPlayback.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            grpPlayback.ColorLighting = Color.FromArgb(29, 200, 238);
            grpPlayback.ColorPen_1 = Color.FromArgb(37, 52, 68);
            grpPlayback.ColorPen_2 = Color.FromArgb(41, 63, 86);
            grpPlayback.Controls.Add(label1);
            grpPlayback.Controls.Add(btnStartPlayback);
            grpPlayback.Controls.Add(btnStopPlayback);
            grpPlayback.Controls.Add(lblLoopCount);
            grpPlayback.Controls.Add(numPlayTimes);
            grpPlayback.CyberGroupBoxStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            grpPlayback.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
            grpPlayback.ForeColor = Color.White;
            grpPlayback.Lighting = false;
            grpPlayback.LinearGradient_Background = false;
            grpPlayback.LinearGradientPen = false;
            grpPlayback.Location = new Point(21, 429);
            grpPlayback.Name = "grpPlayback";
            grpPlayback.PenWidth = 15;
            grpPlayback.RGB = false;
            grpPlayback.Rounding = true;
            grpPlayback.RoundingInt = 12;
            grpPlayback.Size = new Size(350, 80);
            grpPlayback.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            grpPlayback.TabIndex = 23;
            grpPlayback.TabStop = false;
            grpPlayback.Tag = "Cyber";
            grpPlayback.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            grpPlayback.Timer_RGB = 300;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(154, 12);
            label1.Name = "label1";
            label1.Size = new Size(55, 15);
            label1.TabIndex = 14;
            label1.Text = "播放控制";
            // 
            // lblLoopCount
            // 
            lblLoopCount.AutoSize = true;
            lblLoopCount.Font = new Font("Microsoft JhengHei UI", 9F);
            lblLoopCount.ForeColor = Color.White;
            lblLoopCount.Location = new Point(212, 46);
            lblLoopCount.Name = "lblLoopCount";
            lblLoopCount.Size = new Size(58, 15);
            lblLoopCount.TabIndex = 13;
            lblLoopCount.Text = "循環次數:";
            // 
            // grpAdvanced
            // 
            grpAdvanced.Alpha = 20;
            grpAdvanced.BackColor = Color.Transparent;
            grpAdvanced.Background = true;
            grpAdvanced.Background_WidthPen = 3F;
            grpAdvanced.BackgroundPen = true;
            grpAdvanced.ColorBackground = Color.FromArgb(45, 45, 48);
            grpAdvanced.ColorBackground_1 = Color.FromArgb(37, 52, 68);
            grpAdvanced.ColorBackground_2 = Color.FromArgb(41, 63, 86);
            grpAdvanced.ColorBackground_Pen = Color.FromArgb(29, 200, 238);
            grpAdvanced.ColorLighting = Color.FromArgb(29, 200, 238);
            grpAdvanced.ColorPen_1 = Color.FromArgb(37, 52, 68);
            grpAdvanced.ColorPen_2 = Color.FromArgb(41, 63, 86);
            grpAdvanced.Controls.Add(label5);
            grpAdvanced.Controls.Add(btnCustomKeys);
            grpAdvanced.Controls.Add(btnScheduler);
            grpAdvanced.Controls.Add(btnStatistics);
            grpAdvanced.CyberGroupBoxStyle = ReaLTaiizor.Enum.Cyber.StateStyle.Custom;
            grpAdvanced.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
            grpAdvanced.ForeColor = Color.White;
            grpAdvanced.Lighting = false;
            grpAdvanced.LinearGradient_Background = false;
            grpAdvanced.LinearGradientPen = false;
            grpAdvanced.Location = new Point(381, 429);
            grpAdvanced.Name = "grpAdvanced";
            grpAdvanced.PenWidth = 15;
            grpAdvanced.RGB = false;
            grpAdvanced.Rounding = true;
            grpAdvanced.RoundingInt = 12;
            grpAdvanced.Size = new Size(273, 80);
            grpAdvanced.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            grpAdvanced.TabIndex = 27;
            grpAdvanced.TabStop = false;
            grpAdvanced.Tag = "Cyber";
            grpAdvanced.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            grpAdvanced.Timer_RGB = 300;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(113, 12);
            label5.Name = "label5";
            label5.Size = new Size(55, 15);
            label5.TabIndex = 14;
            label5.Text = "進階功能";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(30, 30, 35);
            ClientSize = new Size(676, 596);
            Controls.Add(txtWindowTitle);
            Controls.Add(lblStatus);
            Controls.Add(txtLogDisplay);
            Controls.Add(btnRefreshWindow);
            Controls.Add(btnImportSettings);
            Controls.Add(btnSaveSettings);
            Controls.Add(btnLockWindow);
            Controls.Add(btnHotkeySettings);
            Controls.Add(grpRecording);
            Controls.Add(grpScript);
            Controls.Add(grpEvents);
            Controls.Add(grpPlayback);
            Controls.Add(grpAdvanced);
            Controls.Add(lblWindowStatus);
            Controls.Add(lblRecordingStatus);
            Controls.Add(lblPlaybackStatus);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)numPlayTimes).EndInit();
            grpRecording.ResumeLayout(false);
            grpRecording.PerformLayout();
            grpScript.ResumeLayout(false);
            grpScript.PerformLayout();
            grpEvents.ResumeLayout(false);
            grpEvents.PerformLayout();
            grpPlayback.ResumeLayout(false);
            grpPlayback.PerformLayout();
            grpAdvanced.ResumeLayout(false);
            grpAdvanced.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ReaLTaiizor.Controls.HopeTextBox txtWindowTitle;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ListBox txtLogDisplay;
        private System.Windows.Forms.Timer monitorTimer;
        
        public ReaLTaiizor.Controls.CyberButton btnStartRecording;
        public ReaLTaiizor.Controls.CyberButton btnStopRecording;
        public ReaLTaiizor.Controls.CyberButton btnRefreshWindow;
        public ReaLTaiizor.Controls.CyberButton btnLockWindow;
        public ReaLTaiizor.Controls.CyberButton btnSaveScript;
        public ReaLTaiizor.Controls.CyberButton btnLoadScript;
        public ReaLTaiizor.Controls.CyberButton btnClearEvents;
        public ReaLTaiizor.Controls.CyberButton btnImportSettings;
        public ReaLTaiizor.Controls.CyberButton btnViewEvents;
        public ReaLTaiizor.Controls.CyberButton btnEditEvents;
        public ReaLTaiizor.Controls.CyberButton btnStartPlayback;
        public ReaLTaiizor.Controls.CyberButton btnStopPlayback;
        public System.Windows.Forms.NumericUpDown numPlayTimes;
        public ReaLTaiizor.Controls.CyberButton btnHotkeySettings;
        public ReaLTaiizor.Controls.CyberButton btnCustomKeys;
        public ReaLTaiizor.Controls.CyberButton btnScheduler;
        public ReaLTaiizor.Controls.CyberButton btnStatistics;
        public ReaLTaiizor.Controls.CyberButton btnSaveSettings;
        
        public System.Windows.Forms.Label lblWindowStatus;
        public System.Windows.Forms.Label lblRecordingStatus;
        public System.Windows.Forms.Label lblPlaybackStatus;
        
        // GroupBox 分組框
        private ReaLTaiizor.Controls.CyberGroupBox grpRecording;
        private ReaLTaiizor.Controls.CyberGroupBox grpScript;
        private ReaLTaiizor.Controls.CyberGroupBox grpEvents;
        private ReaLTaiizor.Controls.CyberGroupBox grpPlayback;
        private ReaLTaiizor.Controls.CyberGroupBox grpAdvanced;
        private System.Windows.Forms.Label lblLoopCount;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label1;
        private Label label5;
    }
}
