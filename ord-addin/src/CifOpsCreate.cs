// Creation commands (change the model). Each persists through EditItemFactory and returns what it created so the caller can read it back.
using System;
using System.Collections.Generic;
using System.Text;
using Bentley.CifNET.GeometryModel.SDK;
using Bentley.CifNET.LinearGeometry;
using Bentley.CifNET.SDK;
using Bentley.CifNET.SDK.Edit;
using Bentley.GeometryNET;

namespace WZTC
{
    internal static partial class CifOps
    {
        static ConsensusConnectionEdit Con()
        {
            ConsensusConnectionEdit con = ConsensusConnectionEdit.GetActive();
            if (con == null) throw new InvalidOperationException("no active CIF connection");
            return con;
        }

        static void Persist(object o)
        {
            ConsensusConnectionEdit con = Con();
            con.PersistTransients();
        }

        // "x,y,r;x,y,r;..." (r = curve radius at that PI, 0 = none)
        static List<HorizontalPIPoint> ParsePIs(string s)
        {
            var l = new List<HorizontalPIPoint>();
            foreach (string part in s.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] f = part.Split(',');
                double x = double.Parse(f[0], System.Globalization.CultureInfo.InvariantCulture);
                double y = double.Parse(f[1], System.Globalization.CultureInfo.InvariantCulture);
                double r = f.Length > 2 ? double.Parse(f[2], System.Globalization.CultureInfo.InvariantCulture) : 0.0;
                l.Add(new HorizontalPIPoint(new DPoint3d(x, y, 0.0), r, new TransitionDefinitionInfo(), new TransitionDefinitionInfo()));
            }
            return l;
        }

        static string CreateAlignment(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            string name = Str(args, "name", "");
            List<HorizontalPIPoint> pis = ParsePIs(Str(args, "pis", ""));
            if (pis.Count < 2) throw new InvalidOperationException("pis needs at least 2 points: 'x,y,r;x,y,r'");
            string step = "StartTransientMode";
            try
            {
            Con().StartTransientMode();
            step = "CreateByPIs";
            Alignment a = gm.CreateByPIs(pis);
            if (a == null) throw new InvalidOperationException("CreateByPIs returned null");
            step = "feature";
            string feat = Str(args, "feature", "");
            if (feat.Length > 0) a.SetFeatureDefinition(feat);
            step = "SetName";
            if (name.Length > 0) a.SetName(name);
            step = "AddStationing";
            a.AddStationing(0.0, Num(args, "startStation", 1000.0), false);
            step = "Persist";
            Persist(a);
            return "OK\nkey\tvalue\nname\t" + Safe(() => a.Name) + "\nlength\t" + Safe(() => a.LinearGeometry.Length);
            }
            catch (Exception ex) { throw new InvalidOperationException("step=" + step + ": " + ex.GetType().Name + ": " + ex.Message + " @ " + (ex.StackTrace ?? "").Replace("\n", " | ")); }
        }

        // profile by pass-through points "station,elevation;..." on the alignment picked by name/idx
        static string CreateProfile(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            Con().StartTransientMode();
            var pts = new List<DPoint3d>();
            foreach (string part in Str(args, "points", "").Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] f = part.Split(',');
                pts.Add(new DPoint3d(double.Parse(f[0], System.Globalization.CultureInfo.InvariantCulture),
                                     double.Parse(f[1], System.Globalization.CultureInfo.InvariantCulture), 0.0));
            }
            if (pts.Count < 2) throw new InvalidOperationException("points needs 'station,elev;station,elev;...'");
            ProfileElement pe = ProfileElement.CreateByStartStationEndStationPassThruPoints(pts[0].X, pts[pts.Count - 1].X, pts.ToArray());
            Profile p = a.CreateProfileByProfileElement(pe, true, false);
            if (p == null) throw new InvalidOperationException("CreateProfileByProfileElement returned null");
            Persist(p);
            return "OK\nkey\tvalue\nalignment\t" + Safe(() => a.Name) + "\nprofiles\t" + Safe(() => Count(a.Profiles));
        }

        static string CreateCorridor(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            Con().StartTransientMode();
            Corridor c = a.CreateCorridorByAlignment(Str(args, "corridor", "WZTC_TEST_COR"));
            if (c == null) throw new InvalidOperationException("CreateCorridorByAlignment returned null");
            Persist(c);
            return "OK\nkey\tvalue\ncorridor\t" + Safe(() => c.Name);
        }
    }
}

namespace WZTC
{
    internal static partial class CifOps
    {
        static string DumpAlignments()
        {
            GeometricModel gm = Model();
            var sb = new StringBuilder("OK\nidx\tname\tlength\tsx\tsy\tex\tey\tstartStation\tinterval\n");
            int i = 0;
            foreach (Alignment a in gm.Alignments)
            {
                Alignment al = a;
                sb.Append(i + TAB + Safe(() => al.Name) + TAB + Safe(() => al.LinearGeometry.Length) + TAB + Safe(() => al.LinearGeometry.StartPoint.Coordinates.X) + TAB +
                          Safe(() => al.LinearGeometry.StartPoint.Coordinates.Y) + TAB + Safe(() => al.LinearGeometry.EndPoint.Coordinates.X) + TAB +
                          Safe(() => al.LinearGeometry.EndPoint.Coordinates.Y) + TAB + Safe(() => al.Stationing.StartStation) + TAB + Safe(() => al.IsIntervalAlignment) + "\n");
                i++;
            }
            return sb.ToString().TrimEnd('\n');
        }
    }
}

namespace WZTC
{
    internal static partial class CifOps
    {
        // delete the alignments at the given dump_alignments indexes ("idx":"23,26"); indexes are resolved BEFORE any delete
        static string DeleteAlignments(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            var want = new List<int>();
            foreach (string s in Str(args, "idx", "").Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries)) want.Add(int.Parse(s));
            var targets = new List<Alignment>();
            int i = 0;
            foreach (Alignment a in gm.Alignments) { if (want.Contains(i)) targets.Add(a); i++; }
            ConsensusConnectionEdit con = Con();
            var sb = new StringBuilder("OK\nidx\tresult\n");
            for (int k = 0; k < targets.Count; k++)
            {
                try
                {
                    ConsensusItem item = targets[k];
                    con.Delete(item);
                    sb.Append(want[k] + "\tdeleted\n");
                }
                catch (Exception ex) { sb.Append(want[k] + "\t" + ex.GetType().Name + ": " + ex.Message + "\n"); }
            }
            return sb.ToString().TrimEnd('\n');
        }
    }
}
