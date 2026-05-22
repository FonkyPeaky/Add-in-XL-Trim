namespace XLTrim.Core
{
    public class ScanResult
    {
        public int StylesMaxId         { get; set; }
        public int StylesNodeCount     { get; set; }
        public int DefinedNamedRanges  { get; set; }
        public int InvalidNamedRanges  { get; set; }
        public int HiddenNamedRanges   { get; set; }
    }
}
