using System.Windows;
using System.Windows.Input;

namespace XLTrim.UI
{
    public partial class BackupDialog : GlassWindow
    {
        public bool ShouldBackup { get; private set; }

        public BackupDialog(string filePath)
        {
            InitializeComponent();
            txtFileName.Text = System.IO.Path.GetFileName(filePath);
        }

        private void BackupButton_Click(object sender, RoutedEventArgs e)
        {
            ShouldBackup = true;
            DialogResult = true;
        }

        private void SkipButton_Click(object sender, RoutedEventArgs e)
        {
            ShouldBackup = false;
            DialogResult = true;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
                DragMove();
        }
    }
}
