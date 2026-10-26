// Typed Civil ops: corridor detail, template drops, active profile/surface.
// ORD 2023 CIF SDK (Bentley.CifNET.GeometryModel.SDK + Edit). Distances/elevations are metres.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Bentley.CifNET.GeometryModel.SDK;
using Bentley.CifNET.SDK.Edit;

namespace WZTC
{
    internal static partial class CifOps
    {
        static Corridor PickCorridor(GeometricModel gm, Dictionary<string, object> args)
        {
            string name = Str(args, "corridor", "");
            if (name.Length == 0) name = Str(args, "name", "");
            int idx = (int)Num(args, "corridorIdx", Num(args, "idx", -1));
            int i = 0;
            foreach (Corridor c in gm.Corridors)
            {
                if (name.Length > 0 ? Safe(() => c.Name) == name : i == idx) return c;
                i++;
            }
            throw new InvalidOperationException("corridor not found (corridor='" + name + "' idx=" + idx + ")");
        }

        // READ-ONLY: corridor + each TemplateDrop (stations, interval, template name).
        static string GetCorridorProperties(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Corridor c = PickCorridor(gm, args);
            var sb = new StringBuilder("OK\n");
            sb.Append("key\tvalue\n");
            sb.Append("name\t" + Safe(() => c.Name) + "\n");
            sb.Append("feature\t" + Safe(() => c.FeatureName) + "\n");
            sb.Append("startDistance\t" + Safe(() => c.StartDistance) + "\n");
            sb.Append("endDistance\t" + Safe(() => c.EndDistance) + "\n");
            sb.Append("templateDrops\t" + Safe(() => Count(c.TemplateDrops)) + "\n");
            sb.Append("section\tdrop\n");
            sb.Append("i\tname\tdescription\tstartStation\tstopStation\tstartDistance\tendDistance\tinterval\ttemplate\n");
            int i = 0;
            foreach (TemplateDrop d in c.TemplateDrops)
            {
                TemplateDrop dd = d;
                string tname = "";
                try
                {
                    TemplateDefinition td = dd.GetTemplate();
                    tname = td == null ? "" : Safe(() => td.Name);
                }
                catch (Exception) { tname = "?"; }
                sb.Append(i + TAB + Safe(() => dd.Name) + TAB + Safe(() => dd.Description) + TAB +
                          Safe(() => dd.StartStation) + TAB + Safe(() => dd.StopStation) + TAB +
                          Safe(() => dd.StartDistance) + TAB + Safe(() => dd.EndDistance) + TAB +
                          Safe(() => dd.Interval) + TAB + tname + "\n");
                i++;
            }
            return sb.ToString().TrimEnd('\n');
        }

        // MUTATION: set the alignment's active profile (Boolean SetActiveProfile). Verify with active_profile / list_profiles.
        static string SetActiveProfile(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            Profile p = PickProfile(a, args);
            Con().StartTransientMode();
            bool ok = a.SetActiveProfile(p);
            Persist(p);
            return "OK\nkey\tvalue\nalignment\t" + Safe(() => a.Name) + "\nprofile\t" + Safe(() => p.Name) +
                   "\nsetActiveReturned\t" + ok + "\nactiveProfile\t" + Safe(() => a.ActiveProfile == null ? "" : a.ActiveProfile.Name);
        }

        // READ-ONLY: GeometricModel.ActiveSurface (get-only in ORD 2023 SDK — no SetActiveTerrain here).
        static string GetActiveTerrain(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            SurfaceEntity s = gm.ActiveSurface;
            var sb = new StringBuilder("OK\nkey\tvalue\n");
            sb.Append("hasActive\t" + (s == null ? "N" : "Y") + "\n");
            if (s == null)
            {
                sb.Append("name\t\nfeature\t\nnote\tno ActiveSurface; engineer sets Active Terrain in ORD (SDK has no setter)\n");
                return sb.ToString().TrimEnd('\n');
            }
            sb.Append("name\t" + Safe(() => s.Name) + "\n");
            sb.Append("feature\t" + Safe(() => s.FeatureName) + "\n");
            sb.Append("type\t" + Safe(() => s.GetType().Name) + "\n");
            return sb.ToString().TrimEnd('\n');
        }

        // MUTATION: Corridor.AddTemplateDrop(TemplateDropParameters). Distances metres along corridor.
        // Args: corridor|name|corridorIdx|idx, template (path for FindTemplateByPath), start, end,
        // optional interval (default 5), minStartTransition, minStopTransition, description, lib.
        static string AddTemplateDrop(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Corridor c = PickCorridor(gm, args);
            string path = Str(args, "template", "");
            if (path.Length == 0) throw new InvalidOperationException("template path required (as in list_templates path column)");
            double start = Num(args, "start", double.NaN);
            double end = Num(args, "end", double.NaN);
            if (double.IsNaN(start) || double.IsNaN(end)) throw new InvalidOperationException("start and end (metres along) required");
            if (end <= start) throw new InvalidOperationException("end must be > start");
            double interval = Num(args, "interval", 5.0);
            double minStart = Num(args, "minStartTransition", 0.0);
            double minStop = Num(args, "minStopTransition", 0.0);
            string desc = Str(args, "description", "");
            string lib = Str(args, "lib", "");
            if (lib.Length == 0) lib = TemplateLibrary.GetDefaultTemplateLibraryPath();

            string step = "LoadLibrary";
            try
            {
                TemplateLibrary tl = TemplateLibrary.Load(lib);
                step = "FindTemplateByPath";
                TemplateDefinition td = tl.FindTemplateByPath(path);
                if (td == null) throw new InvalidOperationException("template not found at path '" + path + "' in lib '" + lib + "'");
                step = "TemplateDropParameters";
                var parms = new TemplateDropParameters(td, start, end);
                parms.Interval = interval;
                parms.MinStartTransition = minStart;
                parms.MinStopTransition = minStop;
                if (desc.Length > 0) parms.Desription = desc; // SDK spelling
                int before = Count(c.TemplateDrops);
                step = "StartTransientMode";
                Con().StartTransientMode();
                step = "AddTemplateDrop";
                TemplateDrop drop = c.AddTemplateDrop(parms);
                if (drop == null) throw new InvalidOperationException("AddTemplateDrop returned null");
                step = "Persist";
                Persist(drop);
                int after = Count(c.TemplateDrops);
                return "OK\nkey\tvalue\ncorridor\t" + Safe(() => c.Name) + "\ntemplate\t" + path + "\nlib\t" + lib +
                       "\nstart\t" + start.ToString("R", CultureInfo.InvariantCulture) +
                       "\nend\t" + end.ToString("R", CultureInfo.InvariantCulture) +
                       "\ninterval\t" + interval.ToString("R", CultureInfo.InvariantCulture) +
                       "\ndropsBefore\t" + before + "\ndropsAfter\t" + after +
                       "\ndropName\t" + Safe(() => drop.Name) + "\ndropStartStation\t" + Safe(() => drop.StartStation) +
                       "\ndropStopStation\t" + Safe(() => drop.StopStation);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("step=" + step + ": " + ex.GetType().Name + ": " + ex.Message);
            }
        }
    }
}
