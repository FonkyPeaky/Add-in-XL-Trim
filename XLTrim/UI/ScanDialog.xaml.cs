using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using XLTrim.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLTrim.UI
{
    public partial class ScanDialog : GlassWindow
    {
        private readonly Excel.Workbook _workbook;
        private ScanResult _scanResult;

        public ScanDialog(Excel.Workbook workbook)
        {
            InitializeComponent();
            _workbook = workbook;
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            await Task.Delay(400);
            await RunScanAsync();
        }

        // ── Scan ─────────────────────────────────────────────────────────────

        private async Task RunScanAsync()
        {
            txtSubtitle.Text = "Scanning...";
            ShowPanel(panelScanning);

            // Small delay so the scanning panel paints before we block the UI thread.
            await Task.Delay(400);

            // Excel COM objects require STA — run on the UI thread, not Task.Run.
            _scanResult = WorkbookScanner.Scan(_workbook);

            ShowPanel(panelResults);
            txtSubtitle.Text = "Scan complete — check the results below.";
            btnClean.IsEnabled = true;

            AnimateCards();
        }

        // ── Animations ───────────────────────────────────────────────────────

        private void AnimateCards()
        {
            // Apparition en cascade
            int[] delays = { 0, 100, 200, 300, 400, 500 };
            var cards = new FrameworkElement[] { card1, card2, card3, card4, card5, card6 };
            for (int i = 0; i < cards.Length; i++)
                AnimateCardIn(cards[i], delays[i]);

            // Count-up après apparition des cartes
            AnimateCountUp(numStylesMaxId,    _scanResult.StylesMaxId,       350);
            AnimateCountUp(numStylesCount,    _scanResult.StylesNodeCount,   450);
            AnimateCountUp(numDefinedRanges,  _scanResult.DefinedNamedRanges, 550);
            AnimateCountUp(numInvalidRanges,  _scanResult.InvalidNamedRanges, 650);
            AnimateCountUp(numHiddenRanges,   _scanResult.HiddenNamedRanges,  750);

            // Score santé + couleurs des dots
            Dispatcher.InvokeAsync(async () =>
            {
                await Task.Delay(900);
                UpdateHealthScore();
                UpdateSeverityDots();
            });
        }

        private void AnimateCardIn(FrameworkElement card, int delayMs)
        {
            var ease = new CubicEase { EasingMode = EasingMode.EaseOut };

            card.BeginAnimation(OpacityProperty, new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(500))
            {
                BeginTime = TimeSpan.FromMilliseconds(delayMs),
                EasingFunction = ease
            });

            var tt = (TranslateTransform)card.RenderTransform;
            tt.BeginAnimation(TranslateTransform.YProperty, new DoubleAnimation(16, 0, TimeSpan.FromMilliseconds(500))
            {
                BeginTime = TimeSpan.FromMilliseconds(delayMs),
                EasingFunction = ease
            });
        }

        private void AnimateCountUp(System.Windows.Controls.TextBlock tb, int target, int delayMs)
        {
            var startTime = DateTime.Now.AddMilliseconds(delayMs);
            var duration  = TimeSpan.FromMilliseconds(900);
            var timer     = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };

            timer.Tick += (s, e) =>
            {
                var now = DateTime.Now;
                if (now < startTime) return;

                double elapsed = (now - startTime).TotalMilliseconds;
                if (elapsed >= duration.TotalMilliseconds)
                {
                    tb.Text = target.ToString("N0");
                    timer.Stop();
                    return;
                }

                double t     = elapsed / duration.TotalMilliseconds;
                double eased = 1 - Math.Pow(1 - t, 3);   // cubic ease-out
                tb.Text = ((int)(eased * target)).ToString("N0");
            };

            timer.Start();
        }

        private void UpdateSeverityDots()
        {
            SetDotColor(dotStylesMaxId,    _scanResult.StylesMaxId > 4000    ? "#EF4444" : "#10B981");
            SetDotColor(dotStylesCount,    _scanResult.StylesNodeCount > 4000 ? "#EF4444" : "#10B981");
            SetDotColor(dotDefinedRanges,  _scanResult.InvalidNamedRanges > 0  ? "#F59E0B" : "#10B981");
            SetDotColor(dotInvalidRanges,  _scanResult.InvalidNamedRanges > 0  ? "#EF4444" : "#10B981");
            SetDotColor(dotHiddenRanges,   _scanResult.HiddenNamedRanges > 0   ? "#F59E0B" : "#10B981");
        }

        private static void SetDotColor(System.Windows.Controls.Border dot, string hex)
        {
            var color = (Color)ColorConverter.ConvertFromString(hex);
            dot.Background = new SolidColorBrush(color);
            dot.Effect = new System.Windows.Media.Effects.DropShadowEffect
            {
                Color = color, BlurRadius = 8, ShadowDepth = 0
            };
        }

        private void UpdateHealthScore()
        {
            int score = 100;
            if      (_scanResult.StylesMaxId > 64000)       score -= 40;
            else if (_scanResult.StylesMaxId > 4000)        score -= 20;
            if      (_scanResult.InvalidNamedRanges > 1000) score -= 30;
            else if (_scanResult.InvalidNamedRanges > 0)    score -= 15;
            if      (_scanResult.HiddenNamedRanges > 100)   score -= 10;
            score = Math.Max(0, score);

            string label    = score >= 80 ? "good" : score >= 50 ? "average" : "critical";
            string hexColor = score >= 80 ? "#10B981" : score >= 50 ? "#F59E0B" : "#EF4444";

            txtHealthScore.Text     = score + "%";
            txtHealthLabel.Text     = label;
            txtHealthScore.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hexColor));
        }

        // ── Nettoyage ─────────────────────────────────────────────────────────

        private async void CleanButton_Click(object sender, RoutedEventArgs e)
        {
            btnClean.IsEnabled = false;
            btnClose.IsEnabled = false;   // prevent closing while COM operations run
            ShowPanel(panelCleaning);
            txtSubtitle.Text = "Cleaning in progress...";

            // WPF StrokeDashArray units = multiples of StrokeThickness (6).
            // Ellipse Ø90 → path radius 42 → C = 2π×42 ≈ 263.9 px
            // Round caps (T/2=3 px per end) → effective full dash = 257.9 px → ÷6 ≈ 43.0
            const double RING_CIRC = 43.0;

            // Runs on the UI (STA) thread and forces a render pass so the ring moves
            // visibly during the synchronous COM work.
            void SetProgress(int pct, string status)
            {
                txtProgressPct.Text = $"{pct}%";
                txtCleanStatus.Text = status;
                progressRing.StrokeDashArray = new DoubleCollection { pct / 100.0 * RING_CIRC, 1000 };
                Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, new Action(() => { }));
            }

            int deletedStyles = 0, deletedNames = 0, unhiddenNames = 0;

            // Brief pause so the cleaning panel paints before blocking the UI thread.
            await Task.Delay(100);

            // Excel COM objects require STA — all calls stay on the UI thread.
            SetProgress(5, "Preparing...");

            deletedStyles = StylesCleaner.Clean(_workbook, ratio =>
                SetProgress((int)(10 + ratio * 40), "Cleaning styles..."));

            SetProgress(55, "Removing invalid ranges...");
            deletedNames = NamedRangesCleaner.DeleteInvalid(_workbook);

            SetProgress(80, "Restoring hidden ranges...");
            unhiddenNames = NamedRangesCleaner.UnhideAll(_workbook);

            SetProgress(95, "Saving file...");
            try
            {
                if (!string.IsNullOrEmpty(_workbook.Path))
                    _workbook.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not save workbook:\n{ex.Message}",
                    "XL Trim", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            SetProgress(100, "Done!");
            await Task.Delay(400);

            txtDoneSummary.Text =
                $"{deletedStyles} styles removed\n" +
                $"{deletedNames} invalid named ranges removed\n" +
                $"{unhiddenNames} hidden ranges restored";

            ShowPanel(panelDone);
            txtSubtitle.Text = "File cleaned successfully.";
            btnClose.IsEnabled = true;
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private void ShowPanel(UIElement panel)
        {
            // Stop the spin animation when leaving the scanning panel.
            if (panel != panelScanning)
            {
                try
                {
                    var sb = (System.Windows.Media.Animation.Storyboard)Resources["SpinAnimation"];
                    sb.Stop(spinRing);
                }
                catch { }
            }

            panelScanning.Visibility = Visibility.Collapsed;
            panelResults.Visibility  = Visibility.Collapsed;
            panelCleaning.Visibility = Visibility.Collapsed;
            panelDone.Visibility     = Visibility.Collapsed;
            panel.Visibility         = Visibility.Visible;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed) DragMove();
        }
    }
}
