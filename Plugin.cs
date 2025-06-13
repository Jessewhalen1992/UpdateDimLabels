using System;
using System.IO;
using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Gis.Map.ObjectData;
using OfficeOpenXml;
using System.Text;


// create a short alias so the code reads cleaner
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(UpdateDimLabels.Plugin))]

namespace UpdateDimLabels
{
    public class Plugin : IExtensionApplication
    {
        private static ExcelLookup _company;
        private static ExcelLookup _purpose;

        public void Initialize()
        {
#if NET8_0_OR_GREATER          // Core / .NET 8 path
    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif

            string dllFolder = Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);

            _company = new ExcelLookup(Path.Combine(dllFolder, "CompanyLookup.xlsx"));
            _purpose = new ExcelLookup(Path.Combine(dllFolder, "PurposeLookup.xlsx"));

            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            doc?.Editor.WriteMessage(
                "\nUpdateDimLabels loaded. Run UPDDIM to update dimensions.");
        }

        public void Terminate() { /* nothing to clean up */ }

        [CommandMethod("UPDDIM")]
        public void UpdateAlignedDimensionLabel()
        {
            Editor ed = AcApp.DocumentManager.MdiActiveDocument.Editor;

            // ── 1. pick aligned dimension ────────────────────────────────────────
            PromptEntityOptions peoDim =
                new PromptEntityOptions("\nSelect aligned dimension: ");
            peoDim.SetRejectMessage("\n...must be an *aligned* dimension.");
            peoDim.AddAllowedClass(typeof(AlignedDimension), exactMatch: true);

            PromptEntityResult perDim = ed.GetEntity(peoDim);
            if (perDim.Status != PromptStatus.OK) return;

            // ── 2. pick polyline ─────────────────────────────────────────────────
            PromptEntityOptions peoPl =
                new PromptEntityOptions("\nSelect polyline with Object-Data: ");
            peoPl.SetRejectMessage("\n...must be a *polyline*.");
            peoPl.AddAllowedClass(typeof(Polyline), exactMatch: false);

            PromptEntityResult perPl = ed.GetEntity(peoPl);
            if (perPl.Status != PromptStatus.OK) return;

            // ── 3. update the dimension ─────────────────────────────────────────
            Database db = AcApp.DocumentManager.MdiActiveDocument.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                AlignedDimension dim =
                    (AlignedDimension)tr.GetObject(perDim.ObjectId, OpenMode.ForWrite);
                Entity pl =
                    (Entity)tr.GetObject(perPl.ObjectId, OpenMode.ForRead);

                // read first Object-Data record we can find
                Dictionary<string, string> od =
                    Helpers.ReadFirstOdRecord(pl);

                if (od == null)
                {
                    ed.WriteMessage("\nNo Object-Data found; nothing updated.");
                    tr.Abort();
                    return;
                }

                string dispNum = Helpers.GetOrEmpty(od, "DISP_NUM");
                string company = Helpers.GetOrEmpty(od, "COMPANY");
                string purpcd = Helpers.GetOrEmpty(od, "PURPCD");

                // corrections -----------------------------------------------------
                dispNum = Helpers.InsertSpaceInDispNum(dispNum); // LOC 123456
                company = _company.Lookup(company);              // CNRL (or raw)
                purpcd = _purpose.Lookup(purpcd);               // P/L R/W (or raw)

                // build text override  -------------------------------------------
                string measurement = dim.Measurement.ToString("0.00");
                string newText =
                    company + "\\X" +
                    measurement + " " + purpcd + "\\X" +
                    dispNum;

                dim.DimensionText = newText;   // override
                tr.Commit();
            }
        }

        [CommandMethod("DIMVER")]
        public void DumpAssemblyVersion()
        {
            var asm = typeof(Plugin).Assembly;
            Autodesk.AutoCAD.ApplicationServices.Application
                .ShowAlertDialog($"Loaded from:\n{asm.Location}\n\nVersion {asm.GetName().Version}");
        }
    }
}
