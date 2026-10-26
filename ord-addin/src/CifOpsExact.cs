// Exact PVI profile: straight ProfileLines + exact symmetric ProfileParabolas joined in a ProfileComplex (instead of CreateProfileByVerticalPIs, whose curve inputs do not map).
using System;
using System.Collections.Generic;
using System.Text;
using Bentley.CifNET.GeometryModel.SDK;
using Bentley.CifNET.LinearGeometry;
using Bentley.CifNET.SDK.Edit;
using Bentley.GeometryNET;

namespace WZTC
{
    internal static partial class CifOps
    {
        // "station,elevation,curveLength;..." (metres)
        static void ParsePVIs(string s, List<double> xs, List<double> ys, List<double> ls)
        {
            foreach (string part in s.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] f = part.Split(',');
                if (f.Length != 3) throw new InvalidOperationException("each PVI is 'station,elevation,curveLength'");
                xs.Add(double.Parse(f[0], System.Globalization.CultureInfo.InvariantCulture));
                ys.Add(double.Parse(f[1], System.Globalization.CultureInfo.InvariantCulture));
                ls.Add(double.Parse(f[2], System.Globalization.CultureInfo.InvariantCulture));
            }
            if (xs.Count < 2) throw new InvalidOperationException("need >= 2 PVIs");
        }

        static ProfileComplex BuildExactProfile(List<double> xs, List<double> ys, List<double> ls, StringBuilder log)
        {
            var els = new List<ProfileElement>();
            DPoint3d cur = new DPoint3d(xs[0], ys[0], 0.0);
            int n = xs.Count;
            for (int i = 1; i < n - 1; i++)
            {
                DPoint3d vpi = new DPoint3d(xs[i], ys[i], 0.0);
                double L = ls[i];
                if (L > 0.0)
                {
                    double g1 = (ys[i] - ys[i - 1]) / (xs[i] - xs[i - 1]), g2 = (ys[i + 1] - ys[i]) / (xs[i + 1] - xs[i]);
                    DPoint3d pvc = new DPoint3d(xs[i] - L / 2.0, ys[i] - g1 * L / 2.0, 0.0), pvt = new DPoint3d(xs[i] + L / 2.0, ys[i] + g2 * L / 2.0, 0.0);
                    if (pvc.X - cur.X > 1e-9) els.Add(new ProfileLine(cur, pvc));
                    els.Add(ProfileParabolaConstructor.createParabolaByEndPtsAndVpiPoint(pvc, pvt, vpi));
                    cur = pvt;
                }
                else
                {
                    els.Add(new ProfileLine(cur, vpi));
                    cur = vpi;
                }
            }
            DPoint3d last = new DPoint3d(xs[n - 1], ys[n - 1], 0.0);
            if (last.X - cur.X > 1e-9) els.Add(new ProfileLine(cur, last));
            if (log != null) log.Append("elements\t" + els.Count + "\n");
            return new ProfileComplex(els.ToArray());
        }

        static double ExpectedY(List<double> xs, List<double> ys, List<double> ls, double x)
        {
            int n = xs.Count;
            for (int i = 1; i < n - 1; i++)
            {
                double L = ls[i];
                if (L > 0.0 && x >= xs[i] - L / 2.0 && x <= xs[i] + L / 2.0)
                {
                    double g1 = (ys[i] - ys[i - 1]) / (xs[i] - xs[i - 1]), g2 = (ys[i + 1] - ys[i]) / (xs[i + 1] - xs[i]);
                    double bvc = xs[i] - L / 2.0;
                    return ys[i] - g1 * L / 2.0 + g1 * (x - bvc) + (g2 - g1) / (2.0 * L) * (x - bvc) * (x - bvc);
                }
            }
            for (int i = 0; i < n - 1; i++)
                if (x >= xs[i] - 1e-9 && x <= xs[i + 1] + 1e-9) return ys[i] + (ys[i + 1] - ys[i]) / (xs[i + 1] - xs[i]) * (x - xs[i]);
            return double.NaN;
        }

        // READ-ONLY: build the exact profile in memory (never attached or persisted) and compare it with the maths at sample stations
        static string ProbeExact(Dictionary<string, object> args)
        {
            var xs = new List<double>(); var ys = new List<double>(); var ls = new List<double>();
            ParsePVIs(Str(args, "pvis", ""), xs, ys, ls);
            var sb = new StringBuilder("OK\nkey\tvalue\n");
            ProfileComplex pc = BuildExactProfile(xs, ys, ls, sb);
            sb.Append("range\t" + Safe(() => pc.ProjectedDistRange.Item1 + ".." + pc.ProjectedDistRange.Item2) + "\n");
            sb.Append("subElements\t" + Safe(() => pc.GetSubProfileElements().Length) + "\n");
            double worst = 0.0; double worstX = 0.0; int n = 0;
            for (int i = 0; i < xs.Count - 1; i++)
                for (int j = 0; j <= 20; j++)
                {
                    double x = xs[i] + (xs[i + 1] - xs[i]) * j / 20.0;
                    double want = ExpectedY(xs, ys, ls, x), got;
                    try { got = pc.GetYAtX(x); } catch (Exception ex) { sb.Append("GetYAtX(" + x + ") threw " + ex.GetType().Name + ": " + ex.Message + "\n"); continue; }
                    double d = Math.Abs(got - want); n++;
                    if (d > worst) { worst = d; worstX = x; }
                }
            sb.Append("samples\t" + n + "\nmaxDeviation\t" + worst.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + "\nworstAt\t" + worstX.ToString("R", System.Globalization.CultureInfo.InvariantCulture));
            return sb.ToString();
        }

        // READ-ONLY: which profile does OpenRoads call active, and is it distinguishable from the others?
        static string ActiveProfileInfo(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            var sb = new StringBuilder("OK\nkey\tvalue\n");
            Profile act = null; string err = "";
            try { act = a.ActiveProfile; } catch (Exception ex) { err = ex.GetType().Name + ": " + ex.Message; }
            sb.Append("alignment\t" + Safe(() => a.Name) + "\nactiveIsNull\t" + (act == null ? "Y" : "N") + "\nactiveError\t" + err + "\n");
            string ar = act == null ? "" : Safe(() => act.ProfileGeometry.ProjectedDistRange.Item1 + ".." + act.ProfileGeometry.ProjectedDistRange.Item2);
            sb.Append("activeRange\t" + ar + "\n");
            int i = 0;
            foreach (Profile q in a.Profiles)
            {
                Profile qq = q;
                string r = Safe(() => qq.ProfileGeometry.ProjectedDistRange.Item1 + ".." + qq.ProfileGeometry.ProjectedDistRange.Item2);
                sb.Append("profile" + i + "\tfeature=" + Safe(() => qq.FeatureName) + " featureDef=" + Safe(() => qq.FeatureDefinition == null ? "" : qq.FeatureDefinition.Name) + " name=" + Safe(() => qq.Name) + " range=" + r + " sameObjectAsActive=" + (act != null && Safe(() => qq.Equals(act)) == "True" ? "Y" : "N") + " sameRangeAsActive=" + (act != null && r == ar ? "Y" : "N") + "\n");
                i++;
            }
            return sb.ToString().TrimEnd('\n');
        }

        // READ-ONLY: feature definitions the open file defines (optional regex "match")
        static string ListFeatureDefs(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            string rx = Str(args, "match", "");
            var sb = new StringBuilder("OK\nname\n");
            foreach (FeatureDefinition fd in gm.FeatureDefinitions)
            {
                string n = Safe(() => fd.Name);
                if (rx.Length == 0 || System.Text.RegularExpressions.Regex.IsMatch(n, rx, System.Text.RegularExpressions.RegexOptions.IgnoreCase)) sb.Append(n + "\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        // Gated in Python (workflow_contract.guard_cif_exec: READY token bound to these exact numbers + open QA task): ONE profile from the PVI table with exact parabolic curves, optionally set active.
        static string CreateProfileExact(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            string aname = Safe(() => a.Name);
            var xs = new List<double>(); var ys = new List<double>(); var ls = new List<double>();
            ParsePVIs(Str(args, "pvis", ""), xs, ys, ls);
            ProfileComplex pc = BuildExactProfile(xs, ys, ls, null);
            int before = Count(a.Profiles);
            string step = "StartTransientMode", featRes = "", featRead = "";
            string feat = Str(args, "feature", "");
            try
            {
                Con().StartTransientMode();
                step = "CreateProfileByProfileElement";
                Profile p = a.CreateProfileByProfileElement(pc, Str(args, "active", "1") == "1", false);
                if (p == null) throw new InvalidOperationException("CreateProfileByProfileElement returned null");
                if (feat.Length > 0) { step = "SetFeatureDefinition"; featRes = Safe(() => p.SetFeatureDefinition(feat)); }
                step = "Persist";
                Persist(p);
                featRead = Safe(() => p.FeatureName);
            }
            catch (Exception ex) { throw new InvalidOperationException("step=" + step + ": " + ex.GetType().Name + ": " + ex.Message + " @ " + (ex.StackTrace ?? "").Replace("\n", " | ")); }
            int idx = -1, k = 0;
            foreach (Profile q in a.Profiles)
            {
                var r = q.ProfileGeometry.ProjectedDistRange;
                if (Math.Abs(r.Item1 - xs[0]) < 0.01 && Math.Abs(r.Item2 - xs[xs.Count - 1]) < 0.01) idx = k;
                k++;
            }
            return "OK\nkey\tvalue\nalignment\t" + aname + "\nprofileIdx\t" + idx + "\nprofilesBefore\t" + before + "\nprofilesAfter\t" + Count(a.Profiles) + "\nfeatureRequested\t" + feat + "\nfeatureSetReturned\t" + featRes + "\nfeatureAfterPersist\t" + featRead;
        }
    }
}
