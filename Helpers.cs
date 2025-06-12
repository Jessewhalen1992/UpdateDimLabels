using System.Collections.Generic;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.Gis.Map.ObjectData;

// clear aliases so we avoid name clashes
using MapConstants = Autodesk.Gis.Map.Constants;
using MapTable = Autodesk.Gis.Map.ObjectData.Table;
using Autodesk.Gis.Map;

namespace UpdateDimLabels
{
    internal static class Helpers
    {
        private static readonly Regex _disp =
            new Regex(@"^([A-Z]{3})(\d+)", RegexOptions.IgnoreCase);

        public static string InsertSpaceInDispNum(string value)
        {
            var m = _disp.Match(value);
            return m.Success ? $"{m.Groups[1].Value} {m.Groups[2].Value}" : value;
        }

        public static string GetOrEmpty(Dictionary<string, string> dict, string key)
        {
            string result;
            return dict != null && dict.TryGetValue(key, out result) ? result : "";
        }

        /// Reads the *first* Object‑Data record on <ent>.
        /// Returns a dictionary fieldName → value, or null if none found.
        public static Dictionary<string, string> ReadFirstOdRecord(Entity ent)
        {
            var mapApp = HostMapApplicationServices.Application;
            var tables = mapApp.ActiveProject.ODTables;

            // ---- Try the most generic “get all OD‑records” call ----
            dynamic dynTables = tables;
            dynamic recs;
            try
            {
                // most SDKs expose GetObjectRecords(ObjectId, OpenMode, bool getAllRows)
                recs = dynTables.GetObjectRecords(
                           ent.ObjectId,
                           MapConstants.OpenMode.OpenForRead,
                           true);
            }
            catch
            {
                // fallback: GetObjectRecords(int tableId, ObjectId, OpenMode, bool)
                recs = dynTables.GetObjectRecords(
                           0,
                           ent.ObjectId,
                           MapConstants.OpenMode.OpenForRead,
                           true);
            }

            if (recs == null || recs.Count == 0)
                return null;

            dynamic rec = recs[0];                    // first record is enough
            string tblName = rec.TableName;
            MapTable tbl = tables[tblName];

            var dict = new Dictionary<string, string>();
            int fldCount = tbl.FieldDefinitions.Count;

            for (int i = 0; i < fldCount; i++)
            {
                FieldDefinition fd = tbl.FieldDefinitions[i];
                dict[fd.Name] = rec[i].StrValue;
            }
            return dict;
        }
    }
}
