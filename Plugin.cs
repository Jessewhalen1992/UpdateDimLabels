// Plugin.cs
using System;
using System.IO;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using OfficeOpenXml;
using Autodesk.Gis.Map.ObjectData;

// short alias
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(UpdateDimLabels.Plugin))]

namespace UpdateDimLabels
{
    public class Plugin : IExtensionApplication
    {
        private static ExcelLookup _company;
        private static ExcelLookup _purpose;
        private static string _logFile;

        public void Initialize()
        {
#if NET8_0_OR_GREATER
            System.Text.Encoding.RegisterProvider(
                System.Text.CodePagesEncodingProvider.Instance);
#endif
            string dllFolder = Path.GetDirectoryName(
                                   typeof(Plugin).Assembly.Location);
            _logFile = Path.Combine(dllFolder, "UpdateDimLabels.log");

            try
            {
                // EPPlus 8+ licence
                ExcelPackage.License.SetNonCommercialOrganization(
                    "YourOrganizationName");
                // —or— for personal use:
                // ExcelPackage.License.SetNonCommercialPersonal("YourName");

                // load look‑up tables
                _company = new ExcelLookup(
                               Path.Combine(dllFolder, "CompanyLookup.xlsx"));
                _purpose = new ExcelLookup(
                               Path.Combine(dllFolder, "PurposeLookup.xlsx"));

                AcApp.DocumentManager.MdiActiveDocument?.Editor.WriteMessage(
                    "\nUpdateDimLabels loaded.  Commands: UPDDIM (dimension) and UPM (mtext).");

                File.AppendAllText(_logFile,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} Initialize: SUCCESS\n");
            }
            catch (System.Exception ex)
            {
                File.AppendAllText(_logFile,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} Initialize: ERROR – {ex}\n");
                throw;
            }
        }

        public void Terminate() { }

        // ------------------------------------------------------------------
        // 1) Existing command – aligned dimensions
        // ------------------------------------------------------------------
        [CommandMethod("UPDDIM")]
        public void UpdateAlignedDimensionLabel()
        {
            var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                // pick an aligned dimension
                var peoDim = new PromptEntityOptions("\nSelect aligned dimension: ");
                peoDim.SetRejectMessage("\n…must be an *aligned* dimension.");
                peoDim.AddAllowedClass(typeof(AlignedDimension), exactMatch: true);
                var perDim = ed.GetEntity(peoDim);
                if (perDim.Status != PromptStatus.OK) return;

                // pick a polyline that carries Object‑Data
                var plObjId = PromptPolyline(ed);
                if (plObjId == ObjectId.Null) return;

                using (var tr = HostApplicationServices.WorkingDatabase
                                  .TransactionManager.StartTransaction())
                {
                    var dim = (AlignedDimension)tr.GetObject(
                                  perDim.ObjectId, OpenMode.ForWrite);
                    var pl = (Entity)tr.GetObject(
                                  plObjId, OpenMode.ForRead);

                    if (!TryGetOdValues(pl, out var dispNum, out var company, out var purpcd))
                    {
                        ed.WriteMessage("\nNo Object‑Data found; nothing updated.");
                        tr.Abort();
                        return;
                    }

                    // preserve any manual override text that has *no* “\” codes
                    string measurement;
                    if (!string.IsNullOrWhiteSpace(dim.DimensionText) &&
                        !dim.DimensionText.Contains("\\"))
                    {
                        measurement = dim.DimensionText;
                        ed.WriteMessage("\nUsing manual dimension text: " + measurement);
                    }
                    else
                    {
                        measurement = Helpers.FormatDim(dim.Measurement);
                    }

                    // company\X<measurement> <purpcd>\P<dispNum>
                    dim.DimensionText =
                        company + @"\X" +
                        measurement + " " + purpcd + @"\P" +
                        dispNum;

                    tr.Commit();
                    Log($"UPDDIM success ({dim.ObjectId})");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError in UPDDIM: " + ex.Message);
                Log("UPDDIM error: " + ex);
            }
        }

        // ------------------------------------------------------------------
        // 2) NEW command – mtext
        // ------------------------------------------------------------------
        [CommandMethod("UPM")]
        public void UpdateMTextLabel()
        {
            var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                // pick an MText entity
                var peoTxt = new PromptEntityOptions("\nSelect mtext: ");
                peoTxt.SetRejectMessage("\n…must be *mtext*.");
                peoTxt.AddAllowedClass(typeof(MText), exactMatch: true);
                var perTxt = ed.GetEntity(peoTxt);
                if (perTxt.Status != PromptStatus.OK) return;

                // pick a polyline that carries Object‑Data
                var plObjId = PromptPolyline(ed);
                if (plObjId == ObjectId.Null) return;

                using (var tr = HostApplicationServices.WorkingDatabase
                                  .TransactionManager.StartTransaction())
                {
                    var mtx = (MText)tr.GetObject(
                                  perTxt.ObjectId, OpenMode.ForWrite);
                    var pl = (Entity)tr.GetObject(
                                  plObjId, OpenMode.ForRead);

                    if (!TryGetOdValues(pl, out var dispNum, out var company, out var purpcd))
                    {
                        ed.WriteMessage("\nNo Object‑Data found; nothing updated.");
                        tr.Abort();
                        return;
                    }

                    string middle = mtx.Text.Trim();  // keep user’s original wording
                    if (middle.Length == 0)
                        middle = "<blank>";

                    // company\n<middle> <purpcd>\n<dispNum>
                    mtx.Contents =
                        company + @"\P" +
                        middle + " " + purpcd + @"\P" +
                        dispNum;

                    tr.Commit();
                    Log($"UPM success ({mtx.ObjectId})");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError in UPM: " + ex.Message);
                Log("UPM error: " + ex);
            }
        }

        // ------------------------------------------------------------------
        // helpers shared by both commands
        // ------------------------------------------------------------------

        /// Prompt for a polyline that contains Object‑Data.  
        /// Returns ObjectId.Null if user cancels.
        private static ObjectId PromptPolyline(Editor ed)
        {
            var peoPl = new PromptEntityOptions("\nSelect polyline with Object‑Data: ");
            peoPl.SetRejectMessage("\n…must be a *polyline*.");
            peoPl.AddAllowedClass(typeof(Polyline), exactMatch: false);
            var perPl = ed.GetEntity(peoPl);
            return perPl.Status == PromptStatus.OK ? perPl.ObjectId : ObjectId.Null;
        }

        /// Read OD values and look them up.  
        /// Returns true = success.
        private static bool TryGetOdValues(
            Entity pl,
            out string dispNum,
            out string company,
            out string purpcd)
        {
            dispNum = company = purpcd = "";
            var od = Helpers.ReadFirstOdRecord(pl);
            if (od == null) return false;

            dispNum = Helpers.InsertSpaceInDispNum(
                          Helpers.GetOrEmpty(od, "DISP_NUM"));
            company = _company.Lookup(
                          Helpers.GetOrEmpty(od, "COMPANY"));
            purpcd = _purpose.Lookup(
                          Helpers.GetOrEmpty(od, "PURPCD"));

            return true;
        }

        private static void Log(string text)
        {
            File.AppendAllText(_logFile,
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {text}\n");
        }
    }
}
