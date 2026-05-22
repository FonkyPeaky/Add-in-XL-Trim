using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Office.Core;
using Microsoft.Win32;
using XLTrim.Core;
using XLTrim.UI;

namespace XLTrim.Ribbon
{
    [ComVisible(true)]
    public class XLTrimRibbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("XLTrim.Ribbon.XLTrimRibbon.xml");
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void OnCleanClick(IRibbonControl control)
        {
            var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook == null)
            {
                MessageBox.Show("No workbook is open.", "XL Trim",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (string.IsNullOrEmpty(workbook.Path))
            {
                MessageBox.Show("Please save your file before continuing.", "XL Trim",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // ── Step 1: offer backup ──────────────────────────────────────────
            var backupDialog = new BackupDialog(workbook.FullName);
            if (backupDialog.ShowDialog() != true) return;

            if (backupDialog.ShouldBackup)
            {
                if (!TryCreateBackup(workbook.FullName))
                    return;   // user chose to abort after a save-dialog cancel or error
            }

            // ── Step 2: scan + clean ──────────────────────────────────────────
            var scanDialog = new ScanDialog(workbook);
            scanDialog.ShowDialog();
        }

        // ── Backup helpers ────────────────────────────────────────────────────

        private static bool TryCreateBackup(string sourceFilePath)
        {
            // Build the suggested filename
            string suggested = BackupManager.SuggestedFileName(sourceFilePath);

            // Default directory: Downloads folder
            string downloads = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads");

            // Build the file extension filter
            string ext    = Path.GetExtension(sourceFilePath).TrimStart('.');
            string filter = ext.ToLower() == "xlsx"
                ? "Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*"
                : $"{ext.ToUpper()} file (*.{ext})|*.{ext}|All files (*.*)|*.*";

            var saveDialog = new SaveFileDialog
            {
                Title            = "Save backup as…",
                FileName         = suggested,
                InitialDirectory = Directory.Exists(downloads) ? downloads : workbook_dir(sourceFilePath),
                Filter           = filter,
                OverwritePrompt  = true,
            };

            if (saveDialog.ShowDialog() != true)
            {
                // User cancelled — ask whether to continue without backup
                var answer = MessageBox.Show(
                    "No backup was created.\nDo you want to continue cleaning without a backup?",
                    "XL Trim — No Backup",
                    MessageBoxButton.YesNo, MessageBoxImage.Warning);
                return answer == MessageBoxResult.Yes;
            }

            try
            {
                BackupManager.CreateBackupTo(sourceFilePath, saveDialog.FileName);
                return true;
            }
            catch (Exception ex)
            {
                var answer = MessageBox.Show(
                    $"Unable to create backup:\n{ex.Message}\n\nContinue anyway?",
                    "XL Trim — Backup Error",
                    MessageBoxButton.YesNo, MessageBoxImage.Warning);
                return answer == MessageBoxResult.Yes;
            }
        }

        private static string workbook_dir(string path)
            => Path.GetDirectoryName(path) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        // ── Ribbon resource ───────────────────────────────────────────────────

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            using (var reader = new StreamReader(stream))
                return reader.ReadToEnd();
        }
    }
}
