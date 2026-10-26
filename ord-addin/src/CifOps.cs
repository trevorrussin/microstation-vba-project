// CIF-typed operations. Everything that touches Bentley.CifNET.* lives here (JIT-loads the SDK only when used).
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Bentley.CifNET.GeometryModel.SDK;
using Bentley.CifNET.SDK;
using Bentley.CifNET.SDK.Edit;
using Bentley.GeometryNET;

namespace WZTC
{
    internal static partial class CifOps
    {
        const string TAB = "\t";

        public static string Describe()
        {
            ConsensusConnectionEdit con = ConsensusConnectionEdit.GetActive();
            if (con == null) return "connection=NONE";
            GeometricModel gm = ConsensusConnectionExtensions.GetActiveGeometricModel(con);
            return "connection=OK" + TAB + "geometricModel=" + (gm == null ? "NONE" : "OK");
        }

        static GeometricModel Model()
        {
            ConsensusConnectionEdit con = ConsensusConnectionEdit.GetActive();
            if (con == null) throw new InvalidOperationException("no active CIF connection (open an OpenRoads model first)");
            GeometricModel gm = ConsensusConnectionExtensions.GetActiveGeometricModel(con);
            if (gm == null) throw new InvalidOperationException("the active model has no CIF geometric model (is it a Drawing/Sheet model?)");
            return gm;
        }

        static string Safe(Func<object> f)
        {
            try { object v = f(); return v == null ? "" : Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture).Replace("\t", " ").Replace("\n", " "); }
            catch (Exception) { return "?"; }
        }

        static int Count(IEnumerable e)
        {
            int n = 0;
            if (e != null) foreach (object o in e) n++;
            return n;
        }

        public static string Run(string cmd, Dictionary<string, object> args)
        {
            switch (cmd)
            {
                case "list_alignments": return ListAlignments();
                case "list_corridors": return ListCorridors();
                case "list_terrains": return ListTerrains();
                case "list_sight_visibility": return ListSightVisibility();
                case "list_point_entities_2d": return ListPointEntities2d();
                case "delete_sight_visibility": return DeleteSightVisibility(args);
                case "get_alignment": return GetAlignment(args);
                case "list_profiles": return ListProfiles();
                case "profile_vertices": return ProfileVertices(args);
                case "sample_profile": return SampleProfile(args);
                case "alignment_points": return AlignmentPoints(args);
                case "probe_profile": return ProbeProfile(args);
                case "create_terrain": return CreateTerrain(args);
                case "list_templates": return ListTemplates(args);
                case "list_named_boundaries": return ListNamedBoundaries();
                case "create_named_boundary": return CreateNamedBoundary(args);
                case "list_models": return ListModels();
                case "delete_named_boundaries": return DeleteNamedBoundaries(args);
                case "delete_model": return DeleteModelOp(args);
                case "probe_types": return ProbeTypes(args);
                case "named_views": return NamedViewsOp(args);
                case "create_model": return CreateModel(args);
                case "list_attachments": return ListAttachments();
                case "attach_model": return AttachModel(args);
                case "set_attachment_scale": return SetAttachmentScale(args);
                case "long_section": return LongSection(args);
                case "set_profile_view": return SetProfileView(args);
                case "probe_doccenter_native": return ProbeDocCenterNative();
                case "probe_annotation_types": return ProbeAnnotationTypes(args);
                case "annotate_profile": return AnnotateProfile(args);
                case "session_keyin": return SessionKeyin(args);
                case "list_annotation_aliases": return ListAnnotationAliases();
                case "drape_preview": return DrapePreview(args);
                case "all_models": return AllGeoModels();
                case "profile_from_terrain": return ProfileFromTerrain(args);
                case "profile_boundaries": return ProfileBoundaries(args);
                case "profile_boundary_data": return ProfileBoundaryData(args);
                case "probe_profile2": return ProbeProfile2();
                case "delete_profile": return DeleteProfile(args);
                case "flush_transients": return FlushTransients();
                case "delete_profiles_window": return DeleteProfilesWindow(args);
                case "reflect": return Reflect(args);
                case "create_profile_vpi": return CreateProfileVPI(args);
                case "create_profile_from_pvis": return CreateProfileFromPVIs(args);
                case "probe_exact": return ProbeExact(args);
                case "list_feature_defs": return ListFeatureDefs(args);
                case "active_profile": return ActiveProfileInfo(args);
                case "create_profile_exact": return CreateProfileExact(args);
                case "dump_alignments": return DumpAlignments();
                case "delete_alignments": return DeleteAlignments(args);
                case "create_alignment_cv": return CreateAlignmentCV(args);
                case "create_alignment": return CreateAlignment(args);
                case "create_profile": return CreateProfile(args);
                case "create_corridor": return CreateCorridor(args);
                case "get_corridor_properties": return GetCorridorProperties(args);
                case "set_active_profile": return SetActiveProfile(args);
                case "get_active_terrain": return GetActiveTerrain(args);
                case "add_template_drop": return AddTemplateDrop(args);
                default: return "ERROR" + TAB + "note=unknown CIF command '" + cmd + "'";
            }
        }

        static string ListAlignments()
        {
            GeometricModel gm = Model();
            var sb = new StringBuilder("OK\n");
            sb.Append("name\tfeature\tactiveProfile\tprofiles\tinterval\n");
            int n = 0;
            foreach (Alignment a in gm.Alignments)
            {
                n++;
                sb.Append(Safe(() => a.Name) + TAB + Safe(() => a.FeatureName) + TAB + Safe(() => a.ActiveProfile == null ? "" : a.ActiveProfile.Name) + TAB +
                          Safe(() => Count(a.Profiles)) + TAB + Safe(() => a.IsIntervalAlignment) + "\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        static string ListCorridors()
        {
            GeometricModel gm = Model();
            var sb = new StringBuilder("OK\n");
            sb.Append("name\tfeature\tstartDistance\tendDistance\ttemplateDrops\n");
            foreach (Corridor c in gm.Corridors)
                sb.Append(Safe(() => c.Name) + TAB + Safe(() => c.FeatureName) + TAB + Safe(() => c.StartDistance) + TAB + Safe(() => c.EndDistance) + TAB +
                          Safe(() => Count(c.TemplateDrops)) + "\n");
            return sb.ToString().TrimEnd('\n');
        }

        static string ListTerrains()
        {
            GeometricModel gm = Model();
            var sb = new StringBuilder("OK\n");
            sb.Append("name\tfeature\n");
            foreach (TerrainSurface t in gm.TerrainSurfaces)
                sb.Append(Safe(() => t.Name) + TAB + Safe(() => t.FeatureName) + "\n");
            return sb.ToString().TrimEnd('\n');
        }

        // Sight Visibility analysis draws large yellow/white circles (Draft_Sight_* symbology).
        static string ListSightVisibility()
        {
            ConsensusConnectionEdit con = ConsensusConnectionEdit.GetActive();
            var sb = new StringBuilder("OK\n");
            sb.Append("file\tidx\tdisplayFlag\teye\tobject\tcenterline\n");
            foreach (GeometricModel gm in ConsensusConnectionExtensions.GetAllGeometricModels(con))
            {
                string file = Safe(() => gm.DgnModel.GetDgnFile().GetFileName());
                int i = 0;
                try
                {
                    foreach (SightVisibilitySection s in gm.SightVisibilitySections)
                    {
                        sb.Append(file + TAB + i + TAB + Safe(() => s.SightLineDisplayFlag) + TAB +
                                  Safe(() => s.EyePositionLine == null ? "" : s.EyePositionLine.Name) + TAB +
                                  Safe(() => s.ObjectPositionLine == null ? "" : s.ObjectPositionLine.Name) + TAB +
                                  Safe(() => s.Centerline == null ? "" : s.Centerline.Name) + "\n");
                        i++;
                    }
                }
                catch (Exception) { }
                if (i == 0) sb.Append(file + TAB + "-\t0\t\t\t\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        static string ListPointEntities2d()
        {
            ConsensusConnectionEdit con = ConsensusConnectionEdit.GetActive();
            var sb = new StringBuilder("OK\n");
            sb.Append("file\tidx\tdesc\tx\ty\tz\n");
            foreach (GeometricModel gm in ConsensusConnectionExtensions.GetAllGeometricModels(con))
            {
                string file = Safe(() => gm.DgnModel.GetDgnFile().GetFileName());
                int i = 0;
                try
                {
                    foreach (PointEntity2d p in gm.PointEntities2d)
                    {
                        DPoint3d c = p.Coordinates;
                        sb.Append(file + TAB + i + TAB + Safe(() => p.Description) + TAB +
                                  Safe(() => c.X) + TAB + Safe(() => c.Y) + TAB + Safe(() => c.Z) + "\n");
                        i++;
                    }
                }
                catch (Exception) { }
            }
            return sb.ToString().TrimEnd('\n');
        }

        static string DeleteSightVisibility(Dictionary<string, object> args)
        {
            ConsensusConnectionEdit con = ConsensusConnectionEdit.GetActive();
            var targets = new List<SightVisibilitySection>();
            foreach (GeometricModel gm in ConsensusConnectionExtensions.GetAllGeometricModels(con))
            {
                try { foreach (SightVisibilitySection s in gm.SightVisibilitySections) targets.Add(s); }
                catch (Exception) { }
            }
            int deleted = 0;
            var sb = new StringBuilder("OK\nidx\tresult\n");
            for (int i = 0; i < targets.Count; i++)
            {
                try
                {
                    con.Delete((ConsensusItem)targets[i]);
                    sb.Append(i + "\tdeleted\n");
                    deleted++;
                }
                catch (Exception ex)
                {
                    sb.Append(i + "\t" + ex.GetType().Name + ": " + ex.Message.Replace("\t", " ").Replace("\n", " ") + "\n");
                }
            }
            sb.Append("deleted\t" + deleted);
            return sb.ToString();
        }
    }
}
