using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLTrim.Core
{
    public static class StylesCleaner
    {
        /// <summary>
        /// Removes unused non-built-in styles from the workbook.
        /// Calls <paramref name="onProgress"/> with a ratio 0.0→1.0 during work
        /// (called from the background thread — keep the delegate lightweight).
        /// Returns the number of deleted styles.
        /// </summary>
        public static int Clean(Excel.Workbook workbook, Action<double> onProgress = null)
        {
            // ── Phase 1 : collect styles in use (0 % → 70 %) ──────────────────
            var usedStyles = new HashSet<string>();
            int sheetCount = workbook.Sheets.Count;
            int sheetIdx   = 0;

            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
                sheetIdx++;
                onProgress?.Invoke((sheetIdx - 1.0) / sheetCount * 0.7);

                Excel.Range used = sheet.UsedRange;
                if (used == null) continue;

                foreach (Excel.Range cell in used.Cells)
                {
                    try
                    {
                        object styleObj = cell.Style;

                        if      (styleObj is Excel.Style excelStyle) usedStyles.Add(excelStyle.Name);
                        else if (styleObj is string styleName)        usedStyles.Add(styleName);
                        else
                        {
                            dynamic dyn = styleObj;
                            usedStyles.Add((string)dyn.Name);
                        }
                    }
                    catch { }
                }

                onProgress?.Invoke((double)sheetIdx / sheetCount * 0.7);
            }

            // ── Phase 2 : identify styles to delete ────────────────────────────
            var toDelete = new List<Excel.Style>();
            foreach (Excel.Style style in workbook.Styles)
            {
                try
                {
                    if (!style.BuiltIn && !usedStyles.Contains(style.Name))
                        toDelete.Add(style);
                }
                catch { }
            }

            // ── Phase 3 : delete (70 % → 100 %) ───────────────────────────────
            int total     = Math.Max(1, toDelete.Count);
            int deleteIdx = 0;

            foreach (var style in toDelete)
            {
                deleteIdx++;
                onProgress?.Invoke(0.7 + (double)deleteIdx / total * 0.3);
                try { style.Delete(); }
                catch { }
            }

            onProgress?.Invoke(1.0);
            return toDelete.Count;
        }
    }
}
