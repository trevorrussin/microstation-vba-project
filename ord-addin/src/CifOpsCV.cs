// Alignment from explicit geometry: tangents + circular fillets, built as a CurveVector (exact control of every vertex).
using System;
using System.Collections.Generic;
using System.Text;
using Bentley.CifNET.GeometryModel.SDK;
using Bentley.CifNET.SDK.Edit;
using Bentley.GeometryNET;

namespace WZTC
{
    internal static partial class CifOps
    {
        static string CreateAlignmentCV(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            string name = Str(args, "name", "");
            string feat = Str(args, "feature", "");
            var P = new List<double[]>();   // x, y, r
            foreach (string part in Str(args, "pis", "").Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] f = part.Split(',');
                P.Add(new double[] { double.Parse(f[0], System.Globalization.CultureInfo.InvariantCulture), double.Parse(f[1], System.Globalization.CultureInfo.InvariantCulture),
                                     f.Length > 2 ? double.Parse(f[2], System.Globalization.CultureInfo.InvariantCulture) : 0.0 });
            }
            if (P.Count < 2) throw new InvalidOperationException("pis needs >= 2 points");
            string step = "build";
            try
            {
                CurveVector cv = CurveVector.Create(CurveVector.BoundaryType.Open);
                double cx = P[0][0], cy = P[0][1];   // current pen position
                for (int i = 1; i < P.Count; i++)
                {
                    double r = (i < P.Count - 1) ? P[i][2] : 0.0;
                    if (r <= 0.0)
                    {
                        cv.Add(CurvePrimitive.CreateLine(new DSegment3d(cx, cy, 0.0, P[i][0], P[i][1], 0.0)));
                        cx = P[i][0]; cy = P[i][1];
                        continue;
                    }
                    double ax = P[i - 1][0] - P[i][0], ay = P[i - 1][1] - P[i][1], bx = P[i + 1][0] - P[i][0], by = P[i + 1][1] - P[i][1];
                    double la = Math.Sqrt(ax * ax + ay * ay), lb = Math.Sqrt(bx * bx + by * by);
                    ax /= la; ay /= la; bx /= lb; by /= lb;
                    double dot = Math.Max(-1.0, Math.Min(1.0, ax * bx + ay * by));
                    double inner = Math.Acos(dot);                 // angle between the two legs at the PI
                    double t = r / Math.Tan(inner / 2.0);          // PI -> tangent point
                    double t1x = P[i][0] + ax * t, t1y = P[i][1] + ay * t, t2x = P[i][0] + bx * t, t2y = P[i][1] + by * t;
                    double mx = ax + bx, my = ay + by, lm = Math.Sqrt(mx * mx + my * my);
                    double cd = r / Math.Sin(inner / 2.0);         // PI -> arc centre along the bisector
                    double ccx = P[i][0] + mx / lm * cd, ccy = P[i][1] + my / lm * cd;
                    double a1 = Math.Atan2(t1y - ccy, t1x - ccx), a2 = Math.Atan2(t2y - ccy, t2x - ccx);
                    double sweep = a2 - a1;
                    while (sweep > Math.PI) sweep -= 2 * Math.PI;
                    while (sweep < -Math.PI) sweep += 2 * Math.PI;
                    cv.Add(CurvePrimitive.CreateLine(new DSegment3d(cx, cy, 0.0, t1x, t1y, 0.0)));
                    cv.Add(CurvePrimitive.CreateArc(new DEllipse3d(ccx, ccy, 0.0, r, 0.0, 0.0, 0.0, r, 0.0, Angle.FromRadians(a1), Angle.FromRadians(sweep))));
                    cx = t2x; cy = t2y;
                }
                cv.Add(CurvePrimitive.CreateLine(new DSegment3d(cx, cy, 0.0, P[P.Count - 1][0], P[P.Count - 1][1], 0.0)));

                step = "StartTransientMode";
                ConsensusConnectionEdit con = Con();
                con.StartTransientMode();
                step = "CreateAlignmentByCurveVector";
                Alignment a = gm.CreateAlignmentByCurveVector(cv);
                if (a == null) throw new InvalidOperationException("CreateAlignmentByCurveVector returned null");
                step = "stationing";
                a.AddStationing(0.0, Num(args, "startStation", 1000.0), false);
                step = "feature";
                if (feat.Length > 0) a.SetFeatureDefinition(feat);
                step = "SetName";
                string got = name.Length > 0 ? a.SetName(name) : "";
                step = "PersistTransients";
                con.PersistTransients();
                return "OK\nkey\tvalue\nsetNameReturned\t" + got + "\nname\t" + Safe(() => a.Name) + "\nfeature\t" + Safe(() => a.FeatureName) + "\nlength\t" + Safe(() => a.LinearGeometry.Length);
            }
            catch (Exception ex) { throw new InvalidOperationException("step=" + step + ": " + ex.GetType().Name + ": " + ex.Message + " @ " + (ex.StackTrace ?? "").Replace("\n", " | ")); }
        }
    }
}

namespace WZTC
{
    internal static partial class CifOps
    {
        // READ-ONLY: build a ProfileElement in memory (never attached/persisted) to learn how the factory maps its point array
        static string ProbeProfile(Dictionary<string, object> args)
        {
            var sb = new StringBuilder("OK\nvariant\tvertices\n");
            double s = Num(args, "s", 0), e = Num(args, "e", 755), L = Num(args, "len", 755);
            double[][] d = new double[][] { new double[] { 0, 100 }, new double[] { 300, 104 }, new double[] { 500, 102 }, new double[] { 755, 106 } };
            string[] names = { "xy(s,z0)", "x0z(s,0,e)", "0xz", "yxz(e,s,0)" };
            int only = (int)Num(args, "v", -1);
            for (int v = 0; v < 4; v++)
            {
                if (only >= 0 && v != only) continue;
                var pts = new List<Bentley.GeometryNET.DPoint3d>();
                foreach (double[] p in d)
                {
                    if (v == 0) pts.Add(new Bentley.GeometryNET.DPoint3d(p[0], p[1], 0));
                    else if (v == 1) pts.Add(new Bentley.GeometryNET.DPoint3d(p[0], 0, p[1]));
                    else if (v == 2) pts.Add(new Bentley.GeometryNET.DPoint3d(0, p[0], p[1]));
                    else pts.Add(new Bentley.GeometryNET.DPoint3d(p[1], p[0], 0));
                }
                string r;
                try
                {
                    Bentley.CifNET.LinearGeometry.ProfileElement pe = Bentley.CifNET.LinearGeometry.ProfileElement.CreateByStartStationEndStationPassThruPoints(s, e, pts.ToArray());
                    var t = new StringBuilder();
                    foreach (double x in new double[] { 0, 100, 300, 500, 755, 1000, 1300, 1755 })
                    {
                        double xx = x;
                        t.Append("y(" + xx + ")=" + Safe(() => pe.GetYAtX(xx)) + " ");
                    }
                    t.Append("range=" + Safe(() => pe.ProjectedDistRange.Item1 + ".." + pe.ProjectedDistRange.Item2) + " ");
r = t.ToString() + " len=" + Safe(() => pe.Length);
                }
                catch (Exception ex) { r = ex.GetType().Name + ": " + ex.Message + " @ " + (ex.StackTrace ?? "").Replace("\n", " | "); }
                sb.Append(names[v] + "\t" + r + "\n");
            }
            return sb.ToString().TrimEnd('\n');
        }
    }
}

namespace WZTC
{
    internal static partial class CifOps
    {
        // terrain from a regular grid: x0,y0,nx,ny,dx,dy,z0,slopeX (z = z0 + slopeX*(x-x0))
        static string CreateTerrain(Dictionary<string, object> args)
        {
            double x0 = Num(args, "x0", 0), y0 = Num(args, "y0", 0), dx = Num(args, "dx", 100), dy = Num(args, "dy", 100), z0 = Num(args, "z0", 100), sl = Num(args, "slopeX", 0.0);
            int nx = (int)Num(args, "nx", 5), ny = (int)Num(args, "ny", 5);
            string step = "build";
            try
            {
                var pts = new List<Bentley.GeometryNET.DPoint3d>();
                for (int i = 0; i < nx; i++)
                    for (int j = 0; j < ny; j++)
                        pts.Add(new Bentley.GeometryNET.DPoint3d(x0 + i * dx, y0 + j * dy, z0 + sl * i * dx));
                GeometricModel gm = Model();
                step = "DTM";
                Bentley.TerrainModelNET.DTM dtm = new Bentley.TerrainModelNET.DTM();
                dtm.AddPoints(pts.ToArray());
                step = "StartTransientMode";
                ConsensusConnectionEdit con = Con();
                con.StartTransientMode();
                step = "CreateTerrainSurface";
                TerrainSurface ts = gm.CreateTerrainSurface(dtm);
                if (ts == null) throw new InvalidOperationException("CreateTerrainSurface returned null");
                step = "PersistTransients";
                con.PersistTransients();
                return "OK\nkey\tvalue\npoints\t" + pts.Count + "\nname\t" + Safe(() => ts.Name);
            }
            catch (Exception ex) { throw new InvalidOperationException("step=" + step + ": " + ex.GetType().Name + ": " + ex.Message + " @ " + (ex.StackTrace ?? "").Replace("\n", " | ")); }
        }
    }
}

namespace WZTC
{
    internal static partial class CifOps
    {
        static void WalkTemplates(Category c, string path, StringBuilder sb)
        {
            foreach (Category sub in c.Categories) WalkTemplates(sub, path + "\\" + Safe(() => sub.Name), sb);
            foreach (TemplateDefinition t in c.Templates) sb.Append(path + "\\" + Safe(() => t.Name) + "\n");
        }

        // READ-ONLY: templates in a library (default library when "lib" is empty)
        static string ListTemplates(Dictionary<string, object> args)
        {
            string lib = Str(args, "lib", "");
            if (lib.Length == 0) lib = TemplateLibrary.GetDefaultTemplateLibraryPath();
            TemplateLibrary tl = TemplateLibrary.Load(lib);
            var sb = new StringBuilder("OK\npath\n");
            sb.Append("[library] " + lib + "\n");
            WalkTemplates(tl.RootCategory, "", sb);
            return sb.ToString().TrimEnd('\n');
        }
    }
}
