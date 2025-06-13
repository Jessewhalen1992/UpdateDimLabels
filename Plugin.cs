// Plugin.cs
using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Gis.Map.ObjectData;
using OfficeOpenXml;

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
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
            string dllFolder = Path.GetDirectoryName(typeof(Plugin).Assembly.Location);
            _logFile = Path.Combine(dllFolder, "UpdateDimLabels.log");

            try
            {
                // EPPlus 8+ licensing
                ExcelPackage.License.SetNonCommercialOrganization("YourOrganizationName");
                // –or– for personal use:
                // ExcelPackage.License.SetNonCommercialPersonal("YourName");

                _company = new ExcelLookup(Path.Combine(dllFolder, "CompanyLookup.xlsx"));
                _purpose = new ExcelLookup(Path.Combine(dllFolder, "PurposeLookup.xlsx"));

                AcApp.DocumentManager.MdiActiveDocument
                    .Editor.WriteMessage("\nUpdateDimLabels loaded. Run UPDDIM to update dimensions.");

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

        [CommandMethod("UPDDIM")]
        public void UpdateAlignedDimensionLabel()
        {
            var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                // 1) pick an aligned dimension
                var peoDim = new PromptEntityOptions("\nSelect aligned dimension: ");
                peoDim.SetRejectMessage("\n…must be an *aligned* dimension.");
                peoDim.AddAllowedClass(typeof(AlignedDimension), exactMatch: true);
                var perDim = ed.GetEntity(peoDim);
                if (perDim.Status != PromptStatus.OK) return;

                // 2) pick a polyline
                var peoPl = new PromptEntityOptions("\nSelect polyline with Object-Data: ");
                peoPl.SetRejectMessage("\n…must be a *polyline*.");
                peoPl.AddAllowedClass(typeof(Polyline), exactMatch: false);
                var perPl = ed.GetEntity(peoPl);
                if (perPl.Status != PromptStatus.OK) return;

                using (var tr = AcApp.DocumentManager.MdiActiveDocument.Database
                                  .TransactionManager.StartTransaction())
                {
                    var dim = (AlignedDimension)tr.GetObject(perDim.ObjectId, OpenMode.ForWrite);
                    var pl = (Entity)tr.GetObject(perPl.ObjectId, OpenMode.ForRead);

                    var od = Helpers.ReadFirstOdRecord(pl);
                    if (od == null)
                    {
                        ed.WriteMessage("\nNo Object-Data found; nothing updated.");
                        tr.Abort();
                        return;
                    }

                    // raw OD values
                    string dispNum = Helpers.GetOrEmpty(od, "DISP_NUM");
                    string company = Helpers.GetOrEmpty(od, "COMPANY");
                    string purpcd = Helpers.GetOrEmpty(od, "PURPCD");

                    // clean up
                    dispNum = Helpers.InsertSpaceInDispNum(dispNum);
                    company = _company.Lookup(company);
                    purpcd = _purpose.Lookup(purpcd);

                    // preserve any manual override (text without backslashes)
                    string overrideText = dim.DimensionText;
                    string measurement;
                    if (!string.IsNullOrWhiteSpace(overrideText)
                        && !overrideText.Contains("\\"))
                    {
                        measurement = overrideText;
                        ed.WriteMessage("\nUsing manual dimension text: " + measurement);
                    }
                    else
                    {
                        measurement = Helpers.FormatDim(dim.Measurement);
                    }

                    // build the new label: company\X<measurement> <PURPCD>\P<dispNum>
                    string newText =
                        company + @"\X" +
                        measurement + " " + purpcd + @"\P" +
                        dispNum;

                    dim.DimensionText = newText;
                    tr.Commit();

                    File.AppendAllText(_logFile,
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} UPDDIM success ({dim.ObjectId}): “{newText}”\n");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError in UPDDIM: " + ex.Message);
                File.AppendAllText(_logFile,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} UPDDIM error: {ex}\n");
            }
        }

        [CommandMethod("DIMVER")]
        public void DumpAssemblyVersion()
        {
            var asm = typeof(Plugin).Assembly;
            AcApp.ShowAlertDialog(
                $"Loaded from:\n{asm.Location}\n\nVersion {asm.GetName().Version}");
        }
    }
}
