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
            ((System.ComponentModel.ISupportInitialize)picPreview).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numPlayTimes).BeginInit();
            SuspendLayout();
            // 
            // txtWindowTitle
            // 
            txtWindowTitle.Location = new Point(12, 12);
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
            lblStatus.Size = new Size(68, 15);
            lblStatus.TabIndex = 1;
            lblStatus.Text = "Status: Idle";
            // 
            // picPreview
            // 
            picPreview.BorderStyle = BorderStyle.FixedSingle;
            picPreview.Location = new Point(12, 100);
            picPreview.Name = "picPreview";
            picPreview.Size = new Size(650, 250);
            picPreview.SizeMode = PictureBoxSizeMode.Zoom;
            picPreview.TabIndex = 1;
            picPreview.TabStop = false;
            picPreview.Click += picPreview_Click;
            // 
            // btnStartRecording
            // 
            btnStartRecording.Location = new Point(12, 370);
            btnStartRecording.Name = "btnStartRecording";
            btnStartRecording.Size = new Size(100, 30);
            btnStartRecording.TabIndex = 4;
            btnStartRecording.Text = "開始錄製";
            btnStartRecording.UseVisualStyleBackColor = true;
            // 
            // btnStopRecording
            // 
            btnStopRecording.Enabled = false;
            btnStopRecording.Location = new Point(120, 370);
            btnStopRecording.Name = "btnStopRecording";
            btnStopRecording.Size = new Size(100, 30);
            btnStopRecording.TabIndex = 5;
            btnStopRecording.Text = "停止錄製";
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
            btnSaveScript.Location = new Point(12, 410);
            btnSaveScript.Name = "btnSaveScript";
            btnSaveScript.Size = new Size(100, 30);
            btnSaveScript.TabIndex = 6;
            btnSaveScript.Text = "保存腳本";
            btnSaveScript.UseVisualStyleBackColor = true;
            // 
            // btnLoadScript
            // 
            btnLoadScript.Location = new Point(120, 410);
            btnLoadScript.Name = "btnLoadScript";
            btnLoadScript.Size = new Size(100, 30);
            btnLoadScript.TabIndex = 7;
            btnLoadScript.Text = "載入腳本";
            btnLoadScript.UseVisualStyleBackColor = true;
            // 
            // btnClearEvents
            // 
            btnClearEvents.Location = new Point(228, 410);
            btnClearEvents.Name = "btnClearEvents";
            btnClearEvents.Size = new Size(100, 30);
            btnClearEvents.TabIndex = 8;
            btnClearEvents.Text = "清除事件";
            btnClearEvents.UseVisualStyleBackColor = true;
            // 
            // btnViewEvents
            // 
            btnViewEvents.Location = new Point(336, 410);
            btnViewEvents.Name = "btnViewEvents";
            btnViewEvents.Size = new Size(100, 30);
            btnViewEvents.TabIndex = 9;
            btnViewEvents.Text = "查看事件";
            btnViewEvents.UseVisualStyleBackColor = true;
            // 
            // btnEditEvents
            // 
            btnEditEvents.Location = new Point(444, 410);
            btnEditEvents.Name = "btnEditEvents";
            btnEditEvents.Size = new Size(100, 30);
            btnEditEvents.TabIndex = 10;
            btnEditEvents.Text = "編輯腳本";
            btnEditEvents.UseVisualStyleBackColor = true;
            // 
            // btnStartPlayback
            // 
            btnStartPlayback.Location = new Point(12, 450);
            btnStartPlayback.Name = "btnStartPlayback";
            btnStartPlayback.Size = new Size(100, 30);
            btnStartPlayback.TabIndex = 11;
            btnStartPlayback.Text = "開始播放";
            btnStartPlayback.UseVisualStyleBackColor = true;
            // 
            // btnStopPlayback
            // 
            btnStopPlayback.Enabled = false;
            btnStopPlayback.Location = new Point(120, 450);
            btnStopPlayback.Name = "btnStopPlayback";
            btnStopPlayback.Size = new Size(100, 30);
            btnStopPlayback.TabIndex = 12;
            btnStopPlayback.Text = "停止播放";
            btnStopPlayback.UseVisualStyleBackColor = true;
            // 
            // numPlayTimes
            // 
            numPlayTimes.Location = new Point(233, 455);
            numPlayTimes.Maximum = new decimal(new int[] { 9999, 0, 0, 0 });
            numPlayTimes.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            numPlayTimes.Name = "numPlayTimes";
            numPlayTimes.Size = new Size(80, 23);
            numPlayTimes.TabIndex = 13;
            numPlayTimes.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // lblWindowStatus
            // 
            lblWindowStatus.AutoSize = true;
            lblWindowStatus.Location = new Point(12, 490);
            lblWindowStatus.Name = "lblWindowStatus";
            lblWindowStatus.Size = new Size(106, 15);
            lblWindowStatus.TabIndex = 14;
            lblWindowStatus.Text = "視窗狀態: 尋找中...";
            // 
            // lblRecordingStatus
            // 
            lblRecordingStatus.AutoSize = true;
            lblRecordingStatus.Location = new Point(12, 510);
            lblRecordingStatus.Name = "lblRecordingStatus";
            lblRecordingStatus.Size = new Size(143, 15);
            lblRecordingStatus.TabIndex = 15;
            lblRecordingStatus.Text = "錄製狀態: 就緒 | 事件數: 0";
            // 
            // lblPlaybackStatus
            // 
            lblPlaybackStatus.AutoSize = true;
            lblPlaybackStatus.Location = new Point(12, 530);
            lblPlaybackStatus.Name = "lblPlaybackStatus";
            lblPlaybackStatus.Size = new Size(85, 15);
            lblPlaybackStatus.TabIndex = 16;
            lblPlaybackStatus.Text = "播放狀態: 就緒";
            // 
            // Form1
            // 
            ClientSize = new Size(674, 560);
            Controls.Add(txtWindowTitle);
            Controls.Add(lblStatus);
            Controls.Add(picPreview);
            Controls.Add(btnRefreshWindow);
            Controls.Add(btnLockWindow);
            Controls.Add(btnStartRecording);
            Controls.Add(btnStopRecording);
            Controls.Add(btnSaveScript);
            Controls.Add(btnLoadScript);
            Controls.Add(btnClearEvents);
            Controls.Add(btnViewEvents);
            Controls.Add(btnEditEvents);
            Controls.Add(btnStartPlayback);
            Controls.Add(btnStopPlayback);
            Controls.Add(numPlayTimes);
            Controls.Add(lblWindowStatus);
            Controls.Add(lblRecordingStatus);
            Controls.Add(lblPlaybackStatus);
            Name = "Form1";
            Text = "Maple Macro";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)picPreview).EndInit();
            ((System.ComponentModel.ISupportInitialize)numPlayTimes).EndInit();
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
        
        public System.Windows.Forms.Label lblWindowStatus;
     public System.Windows.Forms.Label lblRecordingStatus;
  public System.Windows.Forms.Label lblPlaybackStatus;
    }
}
