using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

[assembly: CommandClass(typeof(TopoTester.EntryPoint))]

namespace TopoTester
{
    public class EntryPoint
    {
        [CommandMethod("TOPOTESTER")]
        public void RunTopoTester()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                CivilDocument civilDoc = CivilApplication.ActiveDocument;

                var surfaceIds = civilDoc.GetSurfaceIds();
                if (surfaceIds.Count == 0)
                {
                    ed.WriteMessage("\n❌ No surfaces found in this drawing.");
                    return;
                }

                var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId surfaceId in surfaceIds)
                    {
                        TinSurface surface = tr.GetObject(surfaceId, OpenMode.ForRead) as TinSurface;
                        if (surface == null)
                        {
                            ed.WriteMessage("\n⚠️ Skipped non-TIN surface.");
                            continue;
                        }

                        Extents3d? maybeBounds = surface.Bounds;
                        if (!maybeBounds.HasValue)
                        {
                            ed.WriteMessage("\n⚠️ Skipped surface with no bounds.");
                            continue;
                        }

                        Extents3d bounds = maybeBounds.Value;
                        Point3d min = bounds.MinPoint;
                        Point3d max = bounds.MaxPoint;

                        PromptDoubleOptions gridOpts = new PromptDoubleOptions("\nEnter grid interval (use smaller values for more detail)")
                        {
                            DefaultValue = 1.0,
                            AllowNegative = false,
                            AllowZero = false,
                            AllowNone = true
                        };
                        PromptDoubleResult gridResult = ed.GetDouble(gridOpts);
                        if (gridResult.Status != PromptStatus.OK)
                        {
                            ed.WriteMessage("\n❌ Grid sampling canceled.");
                            return;
                        }

                        double interval = gridResult.Value;
                        int width = (int)Math.Ceiling((max.X - min.X) / interval);
                        int height = (int)Math.Ceiling((max.Y - min.Y) / interval);
                        const int maxSize = 5000;

                        if (width > maxSize || height > maxSize)
                        {
                            ed.WriteMessage($"\n❌ Image too large ({width}x{height}). Try increasing grid interval.");
                            return;
                        }

                        string safeName = string.Join("_", surface.Name.Split(Path.GetInvalidFileNameChars()));
                        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        string csvPath = Path.Combine(desktop, $"Topo_{safeName}_{timestamp}.csv");
                        string imagePath = Path.Combine(desktop, $"Topo_{safeName}_{timestamp}.png");

                        List<string> csvLines = new List<string>();
                        Bitmap bmp = new Bitmap(width, height);
                        double zMin = double.MaxValue, zMax = double.MinValue;
                        List<SamplePoint> samples = new List<SamplePoint>();

                        for (double x = min.X; x <= max.X; x += interval)
                        {
                            for (double y = min.Y; y <= max.Y; y += interval)
                            {
                                try
                                {
                                    double z = surface.FindElevationAtXY(x, y);
                                    zMin = Math.Min(zMin, z);
                                    zMax = Math.Max(zMax, z);
                                    samples.Add(new SamplePoint(x, y, z));
                                    csvLines.Add($"{x:F3},{y:F3},{z:F3}");
                                }
                                catch { }
                            }
                        }

                        if (samples.Count == 0)
                        {
                            ed.WriteMessage("\n⚠️ No valid sample points found on surface.");
                            continue;
                        }

                        using (Graphics g = Graphics.FromImage(bmp))
                            g.Clear(Color.Black);

                        foreach (var pt in samples)
                        {
                            double normZ = Clamp((pt.Z - zMin) / (zMax - zMin + 1e-6), 0, 1);
                            Color col = GetHeatColor(normZ);

                            int px = (int)((pt.X - min.X) / interval);
                            int py = height - 1 - (int)((pt.Y - min.Y) / interval);

                            if (px >= 0 && px < width && py >= 0 && py < height)
                                bmp.SetPixel(px, py, col);
                        }

                        File.WriteAllLines(csvPath, csvLines);
                        bmp.Save(imagePath, ImageFormat.Png);
                        stopwatch.Stop();

                        ed.WriteMessage($"\n✅ {samples.Count} points sampled.");
                        ed.WriteMessage($"\n📄 CSV: {csvPath}");
                        ed.WriteMessage($"\n🖼️ Heatmap: {imagePath}");
                        ed.WriteMessage($"\n⏱️ Elapsed time: {stopwatch.Elapsed.TotalSeconds:F2} seconds.");
                        return;
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ Error: {ex.Message}");
            }
        }

        public struct SamplePoint
        {
            public double X, Y, Z;
            public SamplePoint(double x, double y, double z) => (X, Y, Z) = (x, y, z);
        }

        public static double Clamp(double value, double min, double max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        public static Color GetHeatColor(double t)
        {
            t = Clamp(t, 0, 1);
            int r = (int)(255 * t);
            int g = 0;
            int b = (int)(255 * (1 - t));
            return Color.FromArgb(255, r, g, b);
        }
    }
}