using Excel = Microsoft.Office.Interop.Excel;

namespace XLTrim.Core
{
    public static class WorkbookScanner
    {
        public static ScanResult Scan(Excel.Workbook workbook)
        {
            var result = new ScanResult();

            result.StylesNodeCount = workbook.Styles.Count;
            result.StylesMaxId     = result.StylesNodeCount - 1;

            int defined = 0, invalid = 0, hidden = 0;

            foreach (Excel.Name name in workbook.Names)
            {
                defined++;

                // ── Invalid detection ────────────────────────────────────────
                bool isInvalid = false;
                try
                {
                    string refersTo = name.RefersTo?.ToString() ?? string.Empty;

                    if (refersTo.Contains("#REF!")  ||
                        refersTo.Contains("#NAME?") ||
                        string.IsNullOrWhiteSpace(refersTo))
                    {
                        isInvalid = true;
                    }
                    else
                    {
                        // Try resolving the range — throws if the sheet doesn't exist
                        // or the reference is otherwise broken (works even when RefersTo
                        // still shows the original formula before Excel rewrites it).
                        var _ = name.RefersToRange;
                    }
                }
                catch { isInvalid = true; }

                if (isInvalid) invalid++;

                // ── Hidden detection ─────────────────────────────────────────
                try { if (!name.Visible) hidden++; }
                catch { /* name.Visible can throw for some built-in names */ }
            }

            result.DefinedNamedRanges = defined;
            result.InvalidNamedRanges = invalid;
            result.HiddenNamedRanges  = hidden;

            return result;
        }
    }
}
