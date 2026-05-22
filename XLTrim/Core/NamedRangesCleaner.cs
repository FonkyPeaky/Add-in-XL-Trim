using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLTrim.Core
{
    public static class NamedRangesCleaner
    {
        // Matches the sheet name in =SheetName!... or ='Sheet Name'!...
        private static readonly Regex _sheetRef =
            new Regex(@"^=(?:'([^']+)'|([^'!#\s]+))!", RegexOptions.Compiled);

        /// <summary>Supprime les plages nommées avec référence invalide (#REF!, feuille manquante, vide).</summary>
        public static int DeleteInvalid(Excel.Workbook workbook)
        {
            // Build a fast sheet-name lookup — avoids COM exceptions for missing-sheet detection.
            var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Excel.Worksheet ws in workbook.Sheets)
            {
                try { sheetNames.Add(ws.Name); } catch { }
            }

            var toDelete = new List<Excel.Name>();

            foreach (Excel.Name name in workbook.Names)
            {
                bool isInvalid = false;
                try
                {
                    string refersTo = name.RefersTo?.ToString() ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(refersTo) ||
                        refersTo.Contains("#REF!")           ||
                        refersTo.Contains("#NAME?"))
                    {
                        isInvalid = true;
                    }
                    else
                    {
                        // Fast path: if the formula references a sheet by name, check the
                        // HashSet — no COM call, no exception, handles 55 000+ ranges cheaply.
                        var m = _sheetRef.Match(refersTo);
                        if (m.Success)
                        {
                            string sheet = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
                            isInvalid = !sheetNames.Contains(sheet);
                        }
                        else
                        {
                            // Fallback for formulas without a sheet qualifier (e.g. constants,
                            // structured references). Only call RefersToRange as a last resort.
                            try { var _ = name.RefersToRange; }
                            catch { isInvalid = true; }
                        }
                    }
                }
                catch { isInvalid = true; }

                if (isInvalid) toDelete.Add(name);
            }

            foreach (var name in toDelete)
            {
                try { name.Delete(); }
                catch { }
            }

            return toDelete.Count;
        }

        /// <summary>Rend visibles toutes les plages nommées cachées.</summary>
        public static int UnhideAll(Excel.Workbook workbook)
        {
            int count = 0;
            foreach (Excel.Name name in workbook.Names)
            {
                try
                {
                    if (!name.Visible) { name.Visible = true; count++; }
                }
                catch { }
            }
            return count;
        }
    }
}
