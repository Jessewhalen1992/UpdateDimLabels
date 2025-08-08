// Plugin.cs
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Gis.Map.ObjectData;
using OfficeOpenXml;
using System;
using System.IO;
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
#if NET8_0_OR_GREATER
                // EPPlus 8+ licence (for net8 build only)
                ExcelPackage.License.SetNonCommercialOrganization("compass");
                // —or— for personal use:
                // ExcelPackage.License.SetNonCommercialPersonal("YourName");
#endif

                // load look‑up tables
                _company = new ExcelLookup(
                               Path.Combine(dllFolder, "CompanyLookup.xlsx"));
                _purpose = new ExcelLookup(
                               Path.Combine(dllFolder, "PurposeLookup.xlsx"));

                AcApp.DocumentManager.MdiActiveDocument?.Editor.WriteMessage(
                    "\nUpdateDimLabels loaded.  Commands: UPDDIM (dimension), UPM (mtext) and UPDLDR (round).");

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

        // ----------------------------------------------------------------------
        // 1) Aligned-dimension command (works AutoCAD 2014‑2025)
        // ----------------------------------------------------------------------
        [CommandMethod("UPDDIM")]
        [CommandMethod("UPD")]
        public void UpdateAlignedDimensionLabel()
        {
            var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                // 1. Pick an *aligned* dimension
                var peoDim = new PromptEntityOptions("\nSelect aligned dimension: ");
                peoDim.SetRejectMessage("\n…must be an *aligned* dimension.");
                peoDim.AddAllowedClass(typeof(AlignedDimension), exactMatch: true);
                var perDim = ed.GetEntity(peoDim);
                if (perDim.Status != PromptStatus.OK) return;

                // 2. Pick the polyline that carries the Object‑Data
                var plObjId = PromptPolyline(ed);
                if (plObjId == ObjectId.Null) return;

                using (var tr = HostApplicationServices.WorkingDatabase
                                           .TransactionManager.StartTransaction())
                {
                    var dim = (AlignedDimension)tr.GetObject(
                                  perDim.ObjectId, OpenMode.ForWrite);
                    var pl = (Entity)tr.GetObject(
                                  plObjId, OpenMode.ForRead);

                    // 3. Grab DISP_NUM / COMPANY / PURPCD from OD
                    if (!TryGetOdValues(pl, out var dispNum, out var company, out var purpcd))
                    {
                        ed.WriteMessage("\nNo Object‑Data found; nothing updated.");
                        tr.Abort();
                        return;
                    }

                    // 4. Decide what to use as the “measurement” text
                    string measurement;
                    double raw;
                    if (!string.IsNullOrWhiteSpace(dim.DimensionText) &&
                        !dim.DimensionText.Contains("\\") &&
                        double.TryParse(dim.DimensionText, out raw))
                    {
                        // manual text is numeric → round it
                        measurement = Helpers.RoundDimLeader(raw);
                        ed.WriteMessage($"\nRounded manual dimension text: {measurement}");
                    }
                    else
                    {
                        measurement = Helpers.RoundDimLeader(dim.Measurement);
                    }

                    /*  New override text:
                     *  Using \X to drop the baseline after line‑1 and \P for paragraph.
                     */
                    dim.DimensionText =
                        company + @"\X" +
                        measurement + " " + purpcd + @"\P" +
                        dispNum;

                    // 5. Compute text height – API‑safe for 2014‑2025
                    double h;
                    var dstr = (DimStyleTableRecord)tr.GetObject(
                                   dim.DimensionStyle, OpenMode.ForRead);
                    h = dstr.Dimtxt;              // Dimtxt = style text height in all releases
                    if (h < Tolerance.Global.EqualVector) h = 2.5;   // fail‑safe default

                    // 6. Build xDir / yDir without the 2021 “XAxis” property
                    Vector3d xDir = (dim.XLine2Point - dim.XLine1Point).GetNormal();
                    Vector3d yDir = xDir.CrossProduct(dim.Normal).GetNormal(); // in‑plane ⟂

                    // 7. Adjust the text position based on the number of lines.
                    // Split the override text on \X (baseline drop) and \P (paragraph)
                    // codes to determine how many visible lines are present.  For one or
                    // two lines we use ½ × h as before; for three or more lines we add
                    // an extra ½ × h per additional line so the leader arrow remains
                    // centred on the text block.
                    int lineCount = 1;
                    try
                    {
                        var parts = dim.DimensionText.Split(
                            new[] { "\\X", "\\P" }, StringSplitOptions.None);
                        lineCount = parts.Length;
                    }
                    catch
                    {
                        lineCount = 1;
                    }

                    // For 1–2 lines use a factor of 1; for 3+ use (lineCount – 1)
                    double lineFactor = Math.Max(1.0, (lineCount - 1));
                    double offset = h * 0.5 * lineFactor;

                    // Shift text block toward the arrow by the computed offset
                    dim.TextPosition = dim.TextPosition - yDir * offset;

                    tr.Commit();
                    Log($"UPD success ({dim.ObjectId})  →  \"{dim.DimensionText}\"");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError in UPD: " + ex.Message);
                Log("UPD error: " + ex);
            }
        }

        // ----------------------------------------------------------------------
        // 2) Mtext command – update *or* create new label
        //      ‑ Press <Enter> or type S to skip the pick and create a fresh label
        // ----------------------------------------------------------------------
        [CommandMethod("UPM")]
        public void UpdateMTextLabel()
        {
            var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                // ── 1. Pick an existing mtext OR skip ───────────────────────────
                var peoTxt = new PromptEntityOptions(
                    "\nSelect mtext or press <Enter>/type S to *skip* and create new: ");
                peoTxt.SetRejectMessage("\n…must be *mtext*.");
                peoTxt.AddAllowedClass(typeof(MText), exactMatch: true);
                peoTxt.AllowNone = true;
                peoTxt.Keywords.Add("Skip");
                var perTxt = ed.GetEntity(peoTxt);

                bool createNew =
                    perTxt.Status == PromptStatus.None ||
                    (perTxt.Status == PromptStatus.Keyword &&
                     perTxt.StringResult.Equals("Skip", StringComparison.OrdinalIgnoreCase));

                if (!createNew && perTxt.Status != PromptStatus.OK) return;

                // ── 2. Pick the polyline that carries Object‑Data ───────────────
                var plObjId = PromptPolyline(ed);
                if (plObjId == ObjectId.Null) return;

                // If we’re making a new label, ask for insertion point now
                Point3d insPt = Point3d.Origin;
                if (createNew)
                {
                    var ppr = ed.GetPoint("\nSpecify insertion point for new label: ");
                    if (ppr.Status != PromptStatus.OK) return;
                    insPt = ppr.Value;
                }

                using (var tr = HostApplicationServices.WorkingDatabase
                                     .TransactionManager.StartTransaction())
                {
                    var pl = (Entity)tr.GetObject(plObjId, OpenMode.ForRead);

                    // ── 3. Read OD and look‑ups ──────────────────────────────────
                    if (!TryGetOdValues(pl, out var dispNum, out var company, out var purpcd))
                    {
                        ed.WriteMessage("\nNo Object‑Data found; nothing updated.");
                        tr.Abort();
                        return;
                    }

                    // ── 4‑A. UPDATE existing mtext ───────────────────────────────
                    if (!createNew)
                    {
                        var mtx = (MText)tr.GetObject(perTxt.ObjectId, OpenMode.ForWrite);

                        string middle = mtx.Text.Trim();
                        if (middle.Length == 0) middle = "<blank>";

                        // company\n<middle> <purpcd>\n<dispNum>
                        mtx.Contents =
                            company + @"\P" +
                            middle + " " + purpcd + @"\P" +
                            dispNum;

                        tr.Commit();
                        Log($"UPM update ({mtx.ObjectId})");
                        return;
                    }

                    // ── 4‑B. CREATE brand‑new mtext ──────────────────────────────
                    var db = HostApplicationServices.WorkingDatabase;
                    var btr = (BlockTableRecord)tr.GetObject(
                                  db.CurrentSpaceId, OpenMode.ForWrite);

                    // Build contents purely from OD
                    string contents =
                        company + @"\P" +
                        purpcd + @"\P" +
                        dispNum;

                    var newMtx = new MText
                    {
                        Location = insPt,
                        TextHeight = 10,
                        LayerId = db.Clayer,
                        Contents = contents,
                        Attachment = AttachmentPoint.BottomCenter   // tweak if you prefer
                    };

                    // Set text style “80L” if present
                    var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                    if (tst.Has("80L"))
                        newMtx.TextStyleId = tst["80L"];

                    btr.AppendEntity(newMtx);
                    tr.AddNewlyCreatedDBObject(newMtx, true);

                    tr.Commit();
                    Log($"UPM create ({newMtx.ObjectId})");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError in UPM: " + ex.Message);
                Log("UPM error: " + ex);
            }
        }

        // ------------------------------------------------------------------
        // 3) Dimension leader rounding
        // ------------------------------------------------------------------
        [CommandMethod("UPDLDR")]
        public void UpdateDimLeaderValue()
        {
            var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                var peo = new PromptEntityOptions("\nSelect dimension: ");
                peo.SetRejectMessage("\n…must be a dimension.");
                peo.AddAllowedClass(typeof(Dimension), exactMatch: false);
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                using (var tr = HostApplicationServices.WorkingDatabase
                                         .TransactionManager.StartTransaction())
                {
                    var dim = (Dimension)tr.GetObject(per.ObjectId, OpenMode.ForWrite);

                    double currentValue;
                    if (!string.IsNullOrWhiteSpace(dim.DimensionText) &&
                        double.TryParse(dim.DimensionText, out var parsed))
                    {
                        currentValue = parsed;
                    }
                    else
                    {
                        currentValue = dim.Measurement;
                    }

                    string rounded = Helpers.RoundDimLeader(currentValue);
                    dim.DimensionText = rounded;

                    tr.Commit();
                    Log($"UPDLDR ({dim.ObjectId}) → {rounded}");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError in UPDLDR: " + ex.Message);
                Log("UPDLDR error: " + ex);
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
