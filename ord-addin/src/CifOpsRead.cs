// Read-only detail commands: alignment geometry/stationing, profiles, elevations, plan points along an alignment.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Bentley.CifNET.GeometryModel.SDK;
using Bentley.CifNET.LinearGeometry;

namespace WZTC
{
    internal static partial class CifOps
    {
        static string R(double v) { return v.ToString("R", CultureInfo.InvariantCulture); }

        static string Str(Dictionary<string, object> a, string k, string dflt)
        {
            object v;
            return a != null && a.TryGetValue(k, out v) && v != null ? Convert.ToString(v, CultureInfo.InvariantCulture) : dflt;
        }

        static double Num(Dictionary<string, object> a, string k, double dflt)
        {
            object v;
            if (a != null && a.TryGetValue(k, out v) && v != null) return Convert.ToDouble(v, CultureInfo.InvariantCulture);
            return dflt;
        }

        // an alignment is selected by "name" or by its position in list_alignments ("idx")
        static Alignment Pick(GeometricModel gm, Dictionary<string, object> args)
        {
            string name = Str(args, "name", "");
            int idx = (int)Num(args, "idx", -1);
            int i = 0;
            foreach (Alignment a in gm.Alignments)
            {
                if (name.Length > 0 ? a.Name == name : i == idx) return a;
                i++;
            }
            throw new InvalidOperationException("alignment not found (name='" + name + "' idx=" + idx + ")");
        }

        // a profile by "profile" (name) or "profileIdx"; default = the alignment's active profile
        static Profile PickProfile(Alignment a, Dictionary<string, object> args)
        {
            string pn = Str(args, "profile", "");
            int pi = (int)Num(args, "profileIdx", -1);
            string act = a.ActiveProfile == null ? "" : a.ActiveProfile.Name;
            int i = 0;
            foreach (Profile p in a.Profiles)
            {
                if (pn.Length > 0 ? p.Name == pn : (pi >= 0 ? i == pi : (act.Length > 0 && p.Name == act))) return p;
                i++;
            }
            throw new InvalidOperationException("profile not found (profile='" + pn + "' profileIdx=" + pi + "; active='" + act + "')");
        }

        static string GetAlignment(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            LinearElement le = a.LinearGeometry;
            var sb = new StringBuilder("OK\n");
            sb.Append("key\tvalue\n");
            sb.Append("name\t" + Safe(() => a.Name) + "\n");
            sb.Append("feature\t" + Safe(() => a.FeatureName) + "\n");
            sb.Append("length\t" + Safe(() => le.Length) + "\n");
            sb.Append("startStation\t" + Safe(() => a.Stationing.StartStation) + "\n");
            sb.Append("startX\t" + Safe(() => le.StartPoint.Coordinates.X) + "\n");
            sb.Append("startY\t" + Safe(() => le.StartPoint.Coordinates.Y) + "\n");
            sb.Append("endX\t" + Safe(() => le.EndPoint.Coordinates.X) + "\n");
            sb.Append("endY\t" + Safe(() => le.EndPoint.Coordinates.Y) + "\n");
            sb.Append("closed\t" + Safe(() => le.IsClosed) + "\n");
            sb.Append("profiles\t" + Safe(() => Count(a.Profiles)) + "\n");
            sb.Append("activeProfile\t" + Safe(() => a.ActiveProfile == null ? "" : a.ActiveProfile.Name) + "\n");
            sb.Append("stationEquations\t" + Safe(() => Count(a.Stationing.StationEquations)) + "\n");
            return sb.ToString().TrimEnd('\n');
        }

        static string ListProfiles()
        {
            GeometricModel gm = Model();
            var sb = new StringBuilder("OK\n");
            sb.Append("alignment\tprofile\tactive\tx0\tx1\tprojectedLength\tvertices\telev0\telev1\n");
            foreach (Alignment a in gm.Alignments)
            {
                string an = Safe(() => a.Name);
                string act = Safe(() => a.ActiveProfile == null ? "" : a.ActiveProfile.Name);
                foreach (Profile p in a.Profiles)
                {
                    string nm = Safe(() => p.Name);
                    ProfileElement pe = null;
                    try { pe = p.ProfileGeometry; } catch (Exception) { }
                    if (pe == null) { sb.Append(an + "\t" + nm + "\t\t?\t?\t?\t?\t?\t?\n"); continue; }
                    double x0 = 0, x1 = 0;
                    try { x0 = pe.ProjectedDistRange.Item1; x1 = pe.ProjectedDistRange.Item2; } catch (Exception) { }
                    sb.Append(an + "\t" + nm + "\t" + (nm == act ? "Y" : "") + "\t" + R(x0) + "\t" + R(x1) + "\t" + Safe(() => pe.ProjectedLength) + "\t" +
                              Safe(() => pe.LineStringVertices.Length) + "\t" + Safe(() => pe.GetYAtX(x0)) + "\t" + Safe(() => pe.GetYAtX(x1)) + "\n");
                }
            }
            return sb.ToString().TrimEnd('\n');
        }

        static string ProfileVertices(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            ProfileElement pe = PickProfile(Pick(gm, args), args).ProfileGeometry;
            var sb = new StringBuilder("OK\n");
            sb.Append("i\tx\ty\n");
            int i = 0;
            foreach (Bentley.GeometryNET.DPoint3d v in pe.LineStringVertices)
            {
                sb.Append(i + "\t" + R(v.X) + "\t" + R(v.Y) + "\n");
                i++;
            }
            return sb.ToString().TrimEnd('\n');
        }

        // elevations of one profile at given x (distance along, ft): {"x":[...]} or {"x0":..,"x1":..,"step":..}
        static string SampleProfile(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            ProfileElement pe = PickProfile(Pick(gm, args), args).ProfileGeometry;
            var xs = new List<double>();
            object raw;
            if (args.TryGetValue("x", out raw) && raw is ArrayList)
            {
                foreach (object o in (ArrayList)raw) xs.Add(Convert.ToDouble(o, CultureInfo.InvariantCulture));
            }
            else
            {
                double x0 = Num(args, "x0", pe.ProjectedDistRange.Item1), x1 = Num(args, "x1", pe.ProjectedDistRange.Item2), st = Num(args, "step", 50);
                if (st <= 0) throw new InvalidOperationException("step must be > 0");
                for (double x = x0; x <= x1 + 1e-9; x += st) xs.Add(x);
            }
            var sb = new StringBuilder("OK\n");
            sb.Append("x\ty\tslope\tdefined\n");
            foreach (double x in xs)
            {
                bool def = pe.IsDefinedAtX(x);
                double xx = x;
                sb.Append(R(x) + "\t" + (def ? Safe(() => pe.GetYAtX(xx)) : "") + "\t" + (def ? Safe(() => pe.GetSlopeAtX(xx)) : "") + "\t" + (def ? "Y" : "N") + "\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        // plan coordinates along an alignment every "step" ft of distance along ("offset" optional, + = left)
        static string AlignmentPoints(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            LinearElement le = Pick(gm, args).LinearGeometry;
            double step = Num(args, "step", 100), off = Num(args, "offset", 0), len = le.Length;
            if (step <= 0) throw new InvalidOperationException("step must be > 0");
            var sb = new StringBuilder("OK\n");
            sb.Append("distance\tx\ty\tz\n");
            for (double d = 0; ; d += step)
            {
                double dd = Math.Min(d, len);
                LinearPoint lp = le.GetPointAtDistanceOffset(dd, off);
                sb.Append(R(dd) + "\t" + R(lp.Coordinates.X) + "\t" + R(lp.Coordinates.Y) + "\t" + R(lp.Coordinates.Z) + "\n");
                if (dd >= len) break;
            }
            return sb.ToString().TrimEnd('\n');
        }
    }
}
