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
            txtWindowTitle = new TextBox();
            lblStatus = new Label();
            picPreview = new PictureBox();
            monitorTimer = new System.Windows.Forms.Timer(components);
            btnStartRecording = new Button();
            btnStopRecording = new Button();
            btnRefreshWindow = new Button();
            btnLockWindow = new Button();
            btnSaveScript = new Button();
            btnLoadScript = new Button();
            btnClearEvents = new Button();
            btnViewEvents = new Button();
            btnEditEvents = new Button();
            btnStartPlayback = new Button();
            btnStopPlayback = new Button();
            numPlayTimes = new NumericUpDown();
            lblWindowStatus = new Label();
            lblRecordingStatus = new Label();
            lblPlaybackStatus = new Label();
            btnHotkeySettings = new Button();
            grpRecording = new GroupBox();
            grpScript = new GroupBox();
            grpEvents = new GroupBox();
            grpPlayback = new GroupBox();
            lblLoopCount = new Label();
            ((System.ComponentModel.ISupportInitialize)picPreview).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numPlayTimes).BeginInit();
            grpRecording.SuspendLayout();
            grpScript.SuspendLayout();
            grpEvents.SuspendLayout();
            grpPlayback.SuspendLayout();
            SuspendLayout();
            // 
            // txtWindowTitle
            // 
            txtWindowTitle.Location = new Point(12, 16);
            txtWindowTitle.Name = "txtWindowTitle";
            txtWindowTitle.Size = new Size(200, 23);
            txtWindowTitle.TabIndex = 0;
            txtWindowTitle.Text = "MapleStory";
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(12, 50);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(61, 15);
            lblStatus.TabIndex = 1;
            lblStatus.Text = "狀態: 就緒";
            // 
            // picPreview
            // 
            picPreview.BorderStyle = BorderStyle.FixedSingle;
            picPreview.Location = new Point(12, 75);
            picPreview.Name = "picPreview";
            picPreview.Size = new Size(571, 220);
            picPreview.SizeMode = PictureBoxSizeMode.Zoom;
            picPreview.TabIndex = 1;
            picPreview.TabStop = false;
            picPreview.Click += picPreview_Click;
            // 
            // btnStartRecording
            // 
            btnStartRecording.Location = new Point(10, 25);
            btnStartRecording.Name = "btnStartRecording";
            btnStartRecording.Size = new Size(65, 30);
            btnStartRecording.TabIndex = 4;
            btnStartRecording.Text = "開始";
            btnStartRecording.UseVisualStyleBackColor = true;
            // 
            // btnStopRecording
            // 
            btnStopRecording.Enabled = false;
            btnStopRecording.Location = new Point(85, 25);
            btnStopRecording.Name = "btnStopRecording";
            btnStopRecording.Size = new Size(65, 30);
            btnStopRecording.TabIndex = 5;
            btnStopRecording.Text = "停止";
            btnStopRecording.UseVisualStyleBackColor = true;
            // 
            // btnRefreshWindow
            // 
            btnRefreshWindow.Location = new Point(220, 12);
            btnRefreshWindow.Name = "btnRefreshWindow";
            btnRefreshWindow.Size = new Size(80, 30);
            btnRefreshWindow.TabIndex = 2;
            btnRefreshWindow.Text = "重新檢測";
            btnRefreshWindow.UseVisualStyleBackColor = true;
            // 
            // btnLockWindow
            // 
            btnLockWindow.Location = new Point(308, 12);
            btnLockWindow.Name = "btnLockWindow";
            btnLockWindow.Size = new Size(80, 30);
            btnLockWindow.TabIndex = 3;
            btnLockWindow.Text = "手動鎖定";
            btnLockWindow.UseVisualStyleBackColor = true;
            // 
            // btnSaveScript
            // 
            btnSaveScript.Location = new Point(10, 25);
            btnSaveScript.Name = "btnSaveScript";
            btnSaveScript.Size = new Size(68, 30);
            btnSaveScript.TabIndex = 6;
            btnSaveScript.Text = "保存";
            btnSaveScript.UseVisualStyleBackColor = true;
            // 
            // btnLoadScript
            // 
            btnLoadScript.Location = new Point(83, 25);
            btnLoadScript.Name = "btnLoadScript";
            btnLoadScript.Size = new Size(68, 30);
            btnLoadScript.TabIndex = 7;
            btnLoadScript.Text = "載入";
            btnLoadScript.UseVisualStyleBackColor = true;
            // 
            // btnClearEvents
            // 
            btnClearEvents.Location = new Point(85, 25);
            btnClearEvents.Name = "btnClearEvents";
            btnClearEvents.Size = new Size(65, 30);
            btnClearEvents.TabIndex = 8;
            btnClearEvents.Text = "清除";
            btnClearEvents.UseVisualStyleBackColor = true;
            // 
            // btnViewEvents
            // 
            btnViewEvents.Location = new Point(10, 25);
            btnViewEvents.Name = "btnViewEvents";
            btnViewEvents.Size = new Size(65, 30);
            btnViewEvents.TabIndex = 9;
            btnViewEvents.Text = "查看";
            btnViewEvents.UseVisualStyleBackColor = true;
            // 
            // btnEditEvents
            // 
            btnEditEvents.Location = new Point(156, 25);
            btnEditEvents.Name = "btnEditEvents";
            btnEditEvents.Size = new Size(68, 30);
            btnEditEvents.TabIndex = 10;
            btnEditEvents.Text = "編輯";
            btnEditEvents.UseVisualStyleBackColor = true;
            // 
            // btnStartPlayback
            // 
            btnStartPlayback.Location = new Point(10, 25);
            btnStartPlayback.Name = "btnStartPlayback";
            btnStartPlayback.Size = new Size(80, 30);
            btnStartPlayback.TabIndex = 11;
            btnStartPlayback.Text = "▶ 開始";
            btnStartPlayback.UseVisualStyleBackColor = true;
            // 
            // btnStopPlayback
            // 
            btnStopPlayback.Enabled = false;
            btnStopPlayback.Location = new Point(100, 25);
            btnStopPlayback.Name = "btnStopPlayback";
            btnStopPlayback.Size = new Size(80, 30);
            btnStopPlayback.TabIndex = 12;
            btnStopPlayback.Text = "■ 停止";
            btnStopPlayback.UseVisualStyleBackColor = true;
            // 
            // numPlayTimes
            // 
            numPlayTimes.Location = new Point(255, 28);
            numPlayTimes.Maximum = new decimal(new int[] { 9999, 0, 0, 0 });
            numPlayTimes.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            numPlayTimes.Name = "numPlayTimes";
            numPlayTimes.Size = new Size(55, 23);
            numPlayTimes.TabIndex = 13;
            numPlayTimes.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // lblWindowStatus
            // 
            lblWindowStatus.AutoSize = true;
            lblWindowStatus.Location = new Point(12, 465);
            lblWindowStatus.Name = "lblWindowStatus";
            lblWindowStatus.Size = new Size(106, 15);
            lblWindowStatus.TabIndex = 14;
            lblWindowStatus.Text = "視窗狀態: 尋找中...";
            // 
            // lblRecordingStatus
            // 
            lblRecordingStatus.AutoSize = true;
            lblRecordingStatus.Location = new Point(12, 485);
            lblRecordingStatus.Name = "lblRecordingStatus";
            lblRecordingStatus.Size = new Size(143, 15);
            lblRecordingStatus.TabIndex = 15;
            lblRecordingStatus.Text = "錄製狀態: 就緒 | 事件數: 0";
            // 
            // lblPlaybackStatus
            // 
            lblPlaybackStatus.AutoSize = true;
            lblPlaybackStatus.Location = new Point(12, 505);
            lblPlaybackStatus.Name = "lblPlaybackStatus";
            lblPlaybackStatus.Size = new Size(85, 15);
            lblPlaybackStatus.TabIndex = 16;
            lblPlaybackStatus.Text = "播放狀態: 就緒";
            // 
            // btnHotkeySettings
            // 
            btnHotkeySettings.Location = new Point(396, 12);
            btnHotkeySettings.Name = "btnHotkeySettings";
            btnHotkeySettings.Size = new Size(100, 30);
            btnHotkeySettings.TabIndex = 17;
            btnHotkeySettings.Text = "⚙ 熱鍵設定";
            btnHotkeySettings.UseVisualStyleBackColor = true;
            // 
            // grpRecording
            // 
            grpRecording.Controls.Add(btnStartRecording);
            grpRecording.Controls.Add(btnStopRecording);
            grpRecording.Location = new Point(12, 305);
            grpRecording.Name = "grpRecording";
            grpRecording.Size = new Size(160, 70);
            grpRecording.TabIndex = 20;
            grpRecording.TabStop = false;
            grpRecording.Text = "錄製控制";
            // 
            // grpScript
            // 
            grpScript.Controls.Add(btnSaveScript);
            grpScript.Controls.Add(btnLoadScript);
            grpScript.Controls.Add(btnEditEvents);
            grpScript.Location = new Point(180, 305);
            grpScript.Name = "grpScript";
            grpScript.Size = new Size(235, 70);
            grpScript.TabIndex = 21;
            grpScript.TabStop = false;
            grpScript.Text = "腳本管理";
            // 
            // grpEvents
            // 
            grpEvents.Controls.Add(btnViewEvents);
            grpEvents.Controls.Add(btnClearEvents);
            grpEvents.Location = new Point(423, 305);
            grpEvents.Name = "grpEvents";
            grpEvents.Size = new Size(160, 70);
            grpEvents.TabIndex = 22;
            grpEvents.TabStop = false;
            grpEvents.Text = "事件管理";
            // 
            // grpPlayback
            // 
            grpPlayback.Controls.Add(btnStartPlayback);
            grpPlayback.Controls.Add(btnStopPlayback);
            grpPlayback.Controls.Add(lblLoopCount);
            grpPlayback.Controls.Add(numPlayTimes);
            grpPlayback.Location = new Point(12, 385);
            grpPlayback.Name = "grpPlayback";
            grpPlayback.Size = new Size(320, 70);
            grpPlayback.TabIndex = 23;
            grpPlayback.TabStop = false;
            grpPlayback.Text = "播放控制";
            // 
            // lblLoopCount
            // 
            lblLoopCount.AutoSize = true;
            lblLoopCount.Location = new Point(190, 32);
            lblLoopCount.Name = "lblLoopCount";
            lblLoopCount.Size = new Size(58, 15);
            lblLoopCount.TabIndex = 13;
            lblLoopCount.Text = "循環次數:";
            // 
            // Form1
            // 
            ClientSize = new Size(605, 530);
            Controls.Add(txtWindowTitle);
            Controls.Add(lblStatus);
            Controls.Add(picPreview);
            Controls.Add(btnRefreshWindow);
            Controls.Add(btnLockWindow);
            Controls.Add(btnHotkeySettings);
            Controls.Add(grpRecording);
            Controls.Add(grpScript);
            Controls.Add(grpEvents);
            Controls.Add(grpPlayback);
            Controls.Add(lblWindowStatus);
            Controls.Add(lblRecordingStatus);
            Controls.Add(lblPlaybackStatus);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "Form1";
            Text = "Maple Macro - 按鍵錄製播放工具";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)picPreview).EndInit();
            ((System.ComponentModel.ISupportInitialize)numPlayTimes).EndInit();
            grpRecording.ResumeLayout(false);
            grpScript.ResumeLayout(false);
            grpEvents.ResumeLayout(false);
            grpPlayback.ResumeLayout(false);
            grpPlayback.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.TextBox txtWindowTitle;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.PictureBox picPreview;
        private System.Windows.Forms.Timer monitorTimer;
        
        public System.Windows.Forms.Button btnStartRecording;
        public System.Windows.Forms.Button btnStopRecording;
        public System.Windows.Forms.Button btnRefreshWindow;
        public System.Windows.Forms.Button btnLockWindow;
        public System.Windows.Forms.Button btnSaveScript;
        public System.Windows.Forms.Button btnLoadScript;
        public System.Windows.Forms.Button btnClearEvents;
        public System.Windows.Forms.Button btnViewEvents;
        public System.Windows.Forms.Button btnEditEvents;
        public System.Windows.Forms.Button btnStartPlayback;
        public System.Windows.Forms.Button btnStopPlayback;
        public System.Windows.Forms.NumericUpDown numPlayTimes;
        public System.Windows.Forms.Button btnHotkeySettings;
        
        public System.Windows.Forms.Label lblWindowStatus;
        public System.Windows.Forms.Label lblRecordingStatus;
        public System.Windows.Forms.Label lblPlaybackStatus;
        
        // GroupBox 分組框
        private System.Windows.Forms.GroupBox grpRecording;
        private System.Windows.Forms.GroupBox grpScript;
        private System.Windows.Forms.GroupBox grpEvents;
        private System.Windows.Forms.GroupBox grpPlayback;
        private System.Windows.Forms.Label lblLoopCount;
    }
}
