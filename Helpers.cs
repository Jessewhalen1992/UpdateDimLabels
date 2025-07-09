// Helpers.cs
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

        /// <summary>
        /// Format a raw dimension value into a string (two decimals, trailing zeros trimmed).
        /// </summary>
        public static string FormatDim(double measurement)
        {
            return measurement.ToString("0.##");
        }

        /// <summary>
        /// Round <paramref name="measurement"/> to the nearest whole number
        /// unless one of the predefined exception values is closer. The
        /// returned string always has two decimals (e.g. "10.00").
        /// </summary>
        public static string RoundDimLeader(double measurement)
        {
            // candidate: nearest whole number
            double bestValue = System.Math.Round(
                                  measurement,
                                  digits: 0,
                                  System.MidpointRounding.AwayFromZero);
            double bestDiff = System.Math.Abs(measurement - bestValue);

            // list of allowed exception values
            double[] exceptions = new double[]
            {
                10.50, 10.06, 3.05, 4.57, 6.10, 15.24,
                30.18, 30.48, 36.58, 18.29, 9.14, 7.62
            };

            foreach (double ex in exceptions)
            {
                double diff = System.Math.Abs(measurement - ex);
                if (diff < bestDiff)
                {
                    bestValue = ex;
                    bestDiff = diff;
                }
            }

            return bestValue.ToString("0.00");
        }

        /// <summary>
        /// Reads the *first* Object-Data record on <ent>.
        /// Returns a dictionary fieldName → value, or null if none found.
        /// </summary>
        public static Dictionary<string, string> ReadFirstOdRecord(Entity ent)
        {
            var mapApp = HostMapApplicationServices.Application;
            var tables = mapApp.ActiveProject.ODTables;

            // ---- Try the most generic “get all OD-records” call ----
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
