using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace UpdateDimLabels
{
    internal sealed class ExcelLookup
    {
        private readonly Dictionary<string, string> _map =
            new Dictionary<string, string>();

        public ExcelLookup(string workbookPath)
        {
            if (!File.Exists(workbookPath)) return;

            using (ExcelPackage pkg = new ExcelPackage(new FileInfo(workbookPath)))
            {
                var sheet = pkg.Workbook.Worksheets["LIST"]
                            ?? pkg.Workbook.Worksheets[0];

                int rows = sheet.Dimension.End.Row;
                for (int r = 1; r <= rows; r++)
                {
                    string key = sheet.Cells[r, 1].Text.Trim();
                    string val = sheet.Cells[r, 2].Text.Trim();

                    if (key.Length != 0 && !_map.ContainsKey(key))
                        _map.Add(key, val);
                }
            }
        }

        /// returns replacement or original value if not found
        public string Lookup(string value)
        {
            string replacement;
            return _map.TryGetValue(value, out replacement) ? replacement : value;
        }
    }
}
