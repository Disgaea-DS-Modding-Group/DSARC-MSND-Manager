namespace NitroStudio2
{
    partial class StreamPlayer
    {
        private LibVLCSharp.WinForms.VideoView videoView;

        private void InitializeComponent()
        {
            this.videoView = new LibVLCSharp.WinForms.VideoView();
            ((System.ComponentModel.ISupportInitialize)(this.videoView)).BeginInit();
            this.SuspendLayout();
            // 
            // videoView
            // 
            this.videoView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.videoView.BackColor = System.Drawing.Color.Black;
            this.videoView.Location = new System.Drawing.Point(0, 0);
            this.videoView.Name = "videoView";
            this.videoView.Size = new System.Drawing.Size(632, 317);
            this.videoView.TabIndex = 0;
            this.videoView.Text = "videoView";
            this.videoView.MediaPlayer = null;
            // 
            // StreamPlayer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 317);
            this.Controls.Add(this.videoView);
            this.Name = "StreamPlayer";
            this.Text = "Stream Player";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.onClose);
            ((System.ComponentModel.ISupportInitialize)(this.videoView)).EndInit();
            this.ResumeLayout(false);
        }
    }
}