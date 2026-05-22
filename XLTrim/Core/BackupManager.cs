using System;
using System.IO;

namespace XLTrim.Core
{
    public static class BackupManager
    {
        /// <summary>
        /// Returns the suggested backup filename (no path) for a given source file.
        /// </summary>
        public static string SuggestedFileName(string sourceFilePath)
        {
            string name      = Path.GetFileNameWithoutExtension(sourceFilePath);
            string ext       = Path.GetExtension(sourceFilePath);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return $"{name}_backup_{timestamp}{ext}";
        }

        /// <summary>
        /// Copies <paramref name="sourceFilePath"/> to <paramref name="destinationPath"/>.
        /// Throws if the destination already exists or the copy fails.
        /// </summary>
        public static void CreateBackupTo(string sourceFilePath, string destinationPath)
        {
            File.Copy(sourceFilePath, destinationPath, overwrite: false);
        }
    }
}
