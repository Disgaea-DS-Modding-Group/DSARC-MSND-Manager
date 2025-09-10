using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using LibVLCSharp.Shared;
using LibVLCSharp.WinForms;

namespace NitroStudio2
{
    public partial class StreamPlayer : Form
    {
        public string Path;
        public MainWindow MainWindow;

        private LibVLC _libVLC;
        private MediaPlayer _mediaPlayer;

        public StreamPlayer(MainWindow m, string path, string name)
        {
            InitializeComponent();

            Text = "Stream Player - " + name + ".strm";

            Path = path;
            MainWindow = m;

            // Initialize LibVLC
            Core.Initialize();
            _libVLC = new LibVLC();
            _mediaPlayer = new MediaPlayer(_libVLC);

            // Attach player to the VideoView control
            videoView.MediaPlayer = _mediaPlayer;

            // Play the file
            PlayStream(Path);
        }

        private void PlayStream(string path)
        {
            using (var media = new Media(_libVLC, path, FromType.FromPath))
            {
                _mediaPlayer.Play(media);
            }
        }

        private void onClose(object sender, FormClosingEventArgs e)
        {
            // Stop playback
            _mediaPlayer?.Stop();
            _mediaPlayer?.Dispose();
            _libVLC?.Dispose();

            // Delete the file in a background thread
            Thread t = new Thread(DeleteFile);
            t.Start();
        }

        private void DeleteFile()
        {
            try
            {
                if (File.Exists(Path))
                    File.Delete(Path);
            }
            catch {  }

            try
            {
                MainWindow.StreamTempCount--;
            }
            catch {  }
        }
    }
}

