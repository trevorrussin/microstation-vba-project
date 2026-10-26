// Named boundaries (MicroStation platform objects: a shape element + name, grouped). Not CIF-specific.
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Bentley.CifNET.DocCenter.Mstn;
using Bentley.CifNET.GeometryModel.SDK;
using Bentley.CifNET.SDK;
using Bentley.CifNET.SDK.Edit;
using Bentley.DgnPlatformNET;
using Bentley.DgnPlatformNET.Elements;
using Bentley.GeometryNET;
using Bentley.MstnPlatformNET;

namespace WZTC
{
    internal static partial class CifOps
    {
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern IntPtr GetModuleHandle(string lpModuleName);
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern IntPtr LoadLibrary(string lpFileName);
        [DllImport("kernel32", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

        // READ-ONLY: resolve native DocCenter create exports inside the live ORD process.
        // Does NOT call them (WString/RefCountedPtr ABI — a wrong call can crash ORD).
        static string ProbeDocCenterNative()
        {
            string dllName = "Bentley.Cif.DocCenter.4.0.dll";
            IntPtr mod = GetModuleHandle(dllName);
            string loadPath = "";
            if (mod == IntPtr.Zero)
            {
                string dir = System.IO.Path.GetDirectoryName(typeof(Alignment).Assembly.Location);
                loadPath = System.IO.Path.Combine(dir ?? "", dllName);
                if (!System.IO.File.Exists(loadPath))
                {
                    string ord = @"C:\Program Files\Bentley\OpenRoads Designer 2023.00\OpenRoadsDesigner\Cif";
                    loadPath = System.IO.Path.Combine(ord, dllName);
                }
                mod = LoadLibrary(loadPath);
            }
            string[] symbols = new string[] {
                "?BuildProfileNamedBoundary@CifNamedBoundarySDK@NamedBoundarySDK@Cif@Bentley@@SA_NUWString@4@000NN_N1NN11NNNNH@Z",
                "?BuildCifProfileDrawingModelFromBoundary@CifDrawCompositionSDK@NamedBoundarySDK@Cif@Bentley@@SA_NAEAUWString@4@U54@1PEAUDgnModelRef@DgnPlatform@4@H1N1@Z",
                "?BuildCifProfileDrawingModelFromBoundaryAndSeed@CifDrawCompositionSDK@NamedBoundarySDK@Cif@Bentley@@SA_NAEAUWString@4@U54@1PEAUDgnModelRef@DgnPlatform@4@H11N1@Z",
                "?BuildCifProfileDrawingModelFromBoundaryGroup@CifDrawCompositionSDK@NamedBoundarySDK@Cif@Bentley@@SA_NAEAV?$bvector@UWString@Bentley@@V?$BentleyAllocator@UWString@Bentley@@@2@@Bstdcxx@4@UWString@4@PEAUDgnModelRef@DgnPlatform@4@H1N1@Z"
            };
            var sb = new StringBuilder("OK\nkey\tvalue\n");
            sb.Append("module\t" + (mod == IntPtr.Zero ? "NOT_LOADED" : "0x" + mod.ToInt64().ToString("x")) + "\n");
            sb.Append("loadPath\t" + loadPath.Replace("\t", " ") + "\n");
            sb.Append("symbol\taddress\n");
            if (mod == IntPtr.Zero)
            {
                sb.Append("(none)\tGetModuleHandle/LoadLibrary failed err=" + Marshal.GetLastWin32Error() + "\n");
                return sb.ToString().TrimEnd('\n');
            }
            int found = 0;
            foreach (string sym in symbols)
            {
                IntPtr p = GetProcAddress(mod, sym);
                if (p != IntPtr.Zero) found++;
                string shortName = sym.StartsWith("?") ? sym.Substring(1, Math.Min(60, sym.IndexOf('@') > 0 ? sym.IndexOf('@') - 1 : 40)) : sym;
                sb.Append(shortName + "\t" + (p == IntPtr.Zero ? "MISSING" : "0x" + p.ToInt64().ToString("x")) + "\n");
            }
            sb.Append("found\t" + found + "/" + symbols.Length + "\n");
            sb.Append("note\taddresses only — do not invoke without a typed shim (WString ABI)\n");
            return sb.ToString().TrimEnd('\n');
        }

        // READ-ONLY: list managed types/methods in loaded Annotation assemblies that look like Element Annotation.
        static string ProbeAnnotationTypes(Dictionary<string, object> args)
        {
            string filter = Str(args, "filter", "Annotat|DrawingBoundary|Processor");
            var sb = new StringBuilder("OK\nassembly\ttype\tkind\tdetail\n");
            int n = 0;
            foreach (System.Reflection.Assembly a in AppDomain.CurrentDomain.GetAssemblies())
            {
                string an = a.GetName().Name ?? "";
                if (an.IndexOf("Annotation", StringComparison.OrdinalIgnoreCase) < 0) continue;
                Type[] types;
                try { types = a.GetTypes(); } catch (System.Reflection.ReflectionTypeLoadException ex) { types = ex.Types; }
                if (types == null) continue;
                foreach (Type t in types)
                {
                    if (t == null) continue;
                    string tn = t.FullName ?? t.Name;
                    bool typeHit = System.Text.RegularExpressions.Regex.IsMatch(tn, filter, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (typeHit)
                    {
                        sb.Append(an + TAB + tn + TAB + "type" + TAB + (t.IsPublic ? "public" : "nonpublic") + "\n");
                        n++;
                    }
                    System.Reflection.BindingFlags fl = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.DeclaredOnly;
                    try
                    {
                        foreach (System.Reflection.MethodInfo m in t.GetMethods(fl))
                        {
                            if (m.IsSpecialName) continue;
                            if (!System.Text.RegularExpressions.Regex.IsMatch(m.Name, filter, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                                && !typeHit) continue;
                            if (!typeHit && !System.Text.RegularExpressions.Regex.IsMatch(m.Name, "Annotat", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) continue;
                            var ps = new List<string>();
                            foreach (System.Reflection.ParameterInfo q in m.GetParameters()) ps.Add(q.ParameterType.Name);
                            sb.Append(an + TAB + tn + TAB + "method" + TAB + m.ReturnType.Name + " " + m.Name + "(" + string.Join(",", ps.ToArray()) + ")" + "\n");
                            n++;
                            if (n > 200) return sb.ToString().TrimEnd('\n') + "\nnote\ttruncated at 200 rows";
                        }
                    }
                    catch (Exception) { }
                }
            }
            if (n == 0) sb.Append("(none)\t\t\tno Annotation assemblies matched or no Annotate* methods\n");
            sb.Append("count\t" + n + "\n");
            return sb.ToString().TrimEnd('\n');
        }

        static DgnModel ActiveDgnModel() { return Model().DgnModel; }

        // READ-ONLY
        static string ListNamedBoundaries()
        {
            DgnModel dm = ActiveDgnModel();
            var sb = new StringBuilder("OK\ngroup\tname\tdescription\tscale\n");
            foreach (NamedBoundaryGroup g in new NamedBoundaryGroupCollection(dm))
            {
                NamedBoundaryGroup gg = g;
                foreach (NamedBoundary b in new NamedBoundaryCollection(dm, gg))
                {
                    NamedBoundary bb = b;
                    sb.Append(Safe(() => gg.Name) + TAB + Safe(() => bb.Name) + TAB + Safe(() => bb.Description) + TAB + Safe(() => bb.DrawingScale) + "\n");
                }
                if (gg.GetCount() == 0) sb.Append(Safe(() => gg.Name) + TAB + "(empty)\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        // rectangle "x0,y0,x1,y1" (axis-aligned) -> shape -> named boundary inside a new group
        static string CreateNamedBoundary(Dictionary<string, object> args)
        {
            DgnModel dm = ActiveDgnModel();
            string[] f = Str(args, "rect", "").Split(',');
            if (f.Length != 4) throw new InvalidOperationException("rect needs 'x0,y0,x1,y1'");
            double x0 = double.Parse(f[0], System.Globalization.CultureInfo.InvariantCulture), y0 = double.Parse(f[1], System.Globalization.CultureInfo.InvariantCulture),
                   x1 = double.Parse(f[2], System.Globalization.CultureInfo.InvariantCulture), y1 = double.Parse(f[3], System.Globalization.CultureInfo.InvariantCulture);
            string group = Str(args, "group", "WZTC_TEST_NB"), name = Str(args, "name", "WZTC_TEST_NB_1");
            string step = "shape";
            Element added = null;
            try
            {
                double k = Num(args, "uorPerUnit", 10000.0 * 3.280833333);   // input in CIF metres -> UORs (file master unit is ft, 10000 UOR/ft)
                x0 *= k; y0 *= k; x1 *= k; y1 *= k;
                DPoint3d[] pts = new DPoint3d[] { new DPoint3d(x0, y0, 0), new DPoint3d(x1, y0, 0), new DPoint3d(x1, y1, 0), new DPoint3d(x0, y1, 0), new DPoint3d(x0, y0, 0) };
                ShapeElement shp = new ShapeElement(dm, null, pts);
                step = "AddToModel";
                if (shp.AddToModel() != StatusInt.Success) throw new InvalidOperationException("AddToModel failed");
                added = shp;
                step = "new NamedBoundary";
                NamedBoundary nb = new NamedBoundary();
                nb.ModelRef = dm;
                nb.GraphicalElement = shp;
                nb.Name = name;
                nb.Description = Str(args, "description", "");
                step = "Save";
                nb.Save();
                step = "group";
                NamedBoundaryGroup g = NamedBoundaryGroup.FindNamedBoundaryGroup(dm, group);
                if (g == null) g = NamedBoundaryGroup.CreateNamedBoundaryGroup(dm, group, "");
                step = "InsertBoundary";
                g.InsertBoundary(nb);
                g.WriteToFile();
                return "OK\nkey\tvalue\nname\t" + Safe(() => nb.Name) + "\ngroup\t" + Safe(() => g.Name) + "\ncount\t" + g.GetCount();
            }
            catch (Exception ex) { string undo = ""; if (added != null) { try { undo = " (shape removed: " + added.DeleteFromModel() + ")"; } catch (Exception) { undo = " (SHAPE LEFT IN MODEL)"; } } throw new InvalidOperationException("step=" + step + undo + ": " + ex.GetType().Name + ": " + ex.Message + " @ " + (ex.StackTrace ?? "").Replace("\n", " | ")); }
        }
    
        // READ-ONLY: every model in the active file (properties of each index item via reflection)
        static string ListModels()
        {
            DgnFile f = ActiveDgnModel().GetDgnFile();
            var sb = new StringBuilder("OK\nmodel\n");
            foreach (object it in f.GetModelIndexCollection())
            {
                var parts = new List<string>();
                foreach (System.Reflection.PropertyInfo pi in it.GetType().GetProperties())
                {
                    if (pi.GetIndexParameters().Length > 0) continue;
                    string v; try { object o = pi.GetValue(it, null); v = o == null ? "" : Convert.ToString(o, System.Globalization.CultureInfo.InvariantCulture); } catch (Exception) { v = "?"; }
                    parts.Add(pi.Name + "=" + v.Replace("\t", " ").Replace("\n", " "));
                }
                sb.Append(string.Join("; ", parts.ToArray()) + "\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        // new model in the active file: type = Drawing | Sheet | Normal
        static string CreateModel(Dictionary<string, object> args)
        {
            string name = Str(args, "name", "");
            if (name.Length == 0) throw new InvalidOperationException("name required");
            DgnModelType t = (DgnModelType)Enum.Parse(typeof(DgnModelType), Str(args, "type", "Drawing"), true);
            bool is3d = Str(args, "is3d", "false").ToLowerInvariant() == "true";
            DgnFile f = ActiveDgnModel().GetDgnFile();
            DgnModelStatus err;
            DgnModel m = f.CreateNewModel(out err, name, t, is3d, null);
            if (m == null) throw new InvalidOperationException("CreateNewModel failed: " + err);
            return "OK\nkey\tvalue\nname\t" + Safe(() => m.ModelName) + "\ntype\t" + t + "\nstatus\t" + err;
        }

        // delete one model by name from the active file (never the active model; only WZTC_TEST* names, scratch test models)
        static string DeleteModelOp(Dictionary<string, object> args)
        {
            string name = Str(args, "name", "");
            if (!name.StartsWith("WZTC_TEST")) throw new InvalidOperationException("refused: only WZTC_TEST* models");
            DgnModel active = ActiveDgnModel();
            if (Safe(() => active.ModelName) == name) throw new InvalidOperationException("refused: active model");
            DgnFile f = active.GetDgnFile();
            ModelId mid = f.FindModelIdByName(name);
            var sb = new StringBuilder("OK\nkey\tvalue\nname\t" + name + "\n");
            foreach (System.Reflection.MethodInfo mi in typeof(DgnFile).GetMethods())
            {
                if (mi.Name != "DeleteModel") continue;
                var ps = mi.GetParameters();
                sb.Append("method\t" + (ps.Length > 0 ? ps[0].ParameterType.FullName : "none") + "\n");
                if (ps.Length == 1 && ps[0].ParameterType == typeof(DgnModel))
                {
                    DgnModel dm = f.FindLoadedModelById(mid);
                    if (dm == null) { StatusInt e; dm = f.LoadRootModelById(out e, mid); }
                    if (dm == null) throw new InvalidOperationException("could not load " + name);
                    object r2 = mi.Invoke(f, new object[] { dm });
                    sb.Append("result\t" + r2 + "\n");
                    return sb.ToString().TrimEnd('\n');
                }
                if (ps.Length == 1 && ps[0].ParameterType == typeof(ModelId))
                {
                    object r = mi.Invoke(f, new object[] { mid });
                    sb.Append("result\t" + r + "\n");
                    return sb.ToString().TrimEnd('\n');
                }
            }
            throw new InvalidOperationException("no DeleteModel(ModelId) on DgnFile; " + sb);
        }

        // READ-ONLY: reference attachments of every model in the file
        static string ListAttachments()
        {
            DgnFile f = ActiveDgnModel().GetDgnFile();
            var sb = new StringBuilder("OK\nmodel\tattachment\n");
            foreach (object it in f.GetModelIndexCollection())
            {
                object oid = it.GetType().GetProperty("Id").GetValue(it, null);
                string mname = Convert.ToString(it.GetType().GetProperty("Name").GetValue(it, null));
                DgnModel m = null;
                string lerr = "";
                try { ModelId mid = f.FindModelIdByName(mname); m = f.FindLoadedModelById(mid); if (m == null) { StatusInt e; m = f.LoadRootModelById(out e, mid); } } catch (Exception ex) { lerr = ex.GetType().Name + ": " + ex.Message; }
                if (m == null) { sb.Append(mname + TAB + "(could not load)\n"); continue; }
                int n = 0;
                try
                {
                    foreach (DgnAttachment a in m.GetDgnAttachments())
                    {
                        DgnAttachment aa = a; n++;
                        sb.Append(mname + TAB + "model=" + Safe(() => aa.AttachModelName) + "; file=" + Safe(() => aa.AttachFileName) + "; logical=" + Safe(() => aa.LogicalName) +
                                  "; scale=" + Safe(() => aa.DisplayScale) + "; clipped=" + Safe(() => aa.IsClipped) + "; displayed=" + Safe(() => aa.IsDisplayed) +
                                  "; origin=" + Safe(() => { DPoint3d o = aa.GetRefOrigin(); return o.X + "," + o.Y + "," + o.Z; }) + "\n");
                    }
                }
                catch (Exception ex) { sb.Append(mname + TAB + "(attachments: " + ex.GetType().Name + ")\n"); }
                if (n == 0) sb.Append(mname + TAB + "(none)\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        static DgnModel LoadByName(DgnFile f, string name)
        {
            ModelId mid = f.FindModelIdByName(name);
            DgnModel m = f.FindLoadedModelById(mid);
            if (m == null) { StatusInt e; m = f.LoadRootModelById(out e, mid); }
            if (m == null) throw new InvalidOperationException("cannot load model '" + name + "'");
            return m;
        }

        // attach a model of THIS file into another model of this file: target, source, scale (display scale), ox, oy (reference origin, UOR)
        static string AttachModel(Dictionary<string, object> args)
        {
            DgnFile f = ActiveDgnModel().GetDgnFile();
            string tname = Str(args, "target", ""), sname = Str(args, "source", "");
            DgnModel target = LoadByName(f, tname);
            // reuse the moniker of an existing attachment of this same file (the Profile drawing attaches 'Profile' from this file)
            DgnDocumentMoniker mon = null;
            string myFile = f.GetFileName();
            foreach (object it in f.GetModelIndexCollection())
            {
                string mn = Convert.ToString(it.GetType().GetProperty("Name").GetValue(it, null));
                DgnModel mm = LoadByName(f, mn);
                foreach (DgnAttachment a in mm.GetDgnAttachments())
                    if (mon == null && Safe(() => a.AttachFileName) == System.IO.Path.GetFileName(myFile)) mon = a.GetAttachMoniker();
            }
            if (mon == null) throw new InvalidOperationException("no same-file attachment found to copy the moniker from (file=" + myFile + ")");
            string step = "CreateDgnAttachment";
            try
            {
                if (Str(args, "preload", "") == "1") { try { LoadByName(f, sname); } catch (Exception) { } }
                DgnAttachment att = target.CreateDgnAttachment(mon, sname);
                if (att == null) throw new InvalidOperationException("CreateDgnAttachment returned null");
                step = "settings";
                att.DisplayScale = Num(args, "scale", 1.0);
                att.SetRefOrigin(new DPoint3d(Num(args, "ox", 0), Num(args, "oy", 0), 0));
                step = "WriteToModel";
                StatusInt r = att.WriteToModel(false);
                return "OK\nkey\tvalue\ntarget\t" + tname + "\nsource\t" + sname + "\nwrite\t" + r + "\nattachments\t" + Safe(() => Count(target.GetDgnAttachments()));
            }
            catch (Exception ex) { throw new InvalidOperationException("step=" + step + ": " + ex.GetType().Name + ": " + ex.Message + " @ " + (ex.StackTrace ?? "").Replace("\n", " | ")); }
        }

        static string SetAttachmentScale(Dictionary<string, object> args)
        {
            DgnFile f = ActiveDgnModel().GetDgnFile();
            DgnModel target = LoadByName(f, Str(args, "target", ""));
            string src = Str(args, "source", "");
            double sc = Num(args, "scale", 1.0);
            foreach (DgnAttachment a in target.GetDgnAttachments())
            {
                if (Safe(() => a.AttachModelName) != src) continue;
                a.StoredScale = sc;
                a.DisplayScale = sc;
                StatusInt r = a.Rewrite(true, false);
                return "OK\nkey\tvalue\nrewrite\t" + r + "\nstoredScale\t" + Safe(() => a.StoredScale) + "\ndisplayScale\t" + Safe(() => a.DisplayScale);
            }
            throw new InvalidOperationException("attachment of '" + src + "' not found in '" + Str(args, "target", "") + "'");
        }

        // READ-ONLY: the long-section (profile) space of an alignment and the DGN model behind it
        static string LongSection(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            var sb = new StringBuilder("OK\nkey\tvalue\n");
            sb.Append("alignment\t" + Safe(() => a.Name) + "\n");
            LongSectionSpace ls = null;
            try { ls = a.LongSectionSpace; } catch (Exception ex) { sb.Append("LongSectionSpace\tERR " + ex.GetType().Name + ": " + ex.Message + "\n"); }
            sb.Append("space\t" + (ls == null ? "null" : "present") + "\n");
            if (ls != null)
            {
                sb.Append("modelName\t" + Safe(() => ls.DgnModel.ModelName) + "\n");
                sb.Append("modelId\t" + Safe(() => ls.DgnModel.GetModelId().ToString()) + "\n");
                sb.Append("modelFile\t" + Safe(() => ls.DgnModel.GetDgnFile().GetFileName()) + "\n");
                sb.Append("isDictionary\t" + Safe(() => ls.DgnModel.IsDictionaryModel) + "\n");
                sb.Append("alignmentOfSpace\t" + Safe(() => ls.Alignment.Name) + "\n");
                sb.Append("entities\t" + Safe(() => Count(ls.FeaturizedModelEntities)) + "\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        // In-process key-in (Session.Keyin / EnqueueKeyin). Outside CadInputQueue often fails to arm Civil
        // DocCenter tools (Place Named Boundary Civil Profile, Annotate Element); this runs inside ORD.
        // mode: keyin (default) | enqueue
        static string SessionKeyin(Dictionary<string, object> args)
        {
            string keyin = Str(args, "keyin", "");
            if (keyin.Length == 0) keyin = Str(args, "cmd", "");
            if (keyin.Length == 0) throw new InvalidOperationException("keyin required");
            string mode = Str(args, "mode", "keyin");
            string before = "";
            try { before = Session.Instance.GetCurrentCommandName() ?? ""; } catch (Exception) { }
            string step = mode;
            try
            {
                if (mode == "enqueue") Session.Instance.EnqueueKeyin(keyin);
                else Session.Instance.Keyin(keyin);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("step=" + step + ": " + ex.GetType().Name + ": " + ex.Message);
            }
            string after = "";
            try { after = Session.Instance.GetCurrentCommandName() ?? ""; } catch (Exception) { }
            return "OK\nkey\tvalue\nkeyin\t" + keyin.Replace("\t", " ") + "\nmode\t" + mode +
                   "\ncommandBefore\t" + before.Replace("\t", " ") + "\ncommandAfter\t" + after.Replace("\t", " ");
        }

        // N1 (Munoz PDF / profile-cutting-workflow): put the alignment's profile model into a view (Open Profile Model).
        // viewIndex is 0-based for LongSectionViewManager (View 1 => 0, View 2 => 1).
        static string SetProfileView(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            LongSectionSpace ls = a.LongSectionSpace;
            if (ls == null) throw new InvalidOperationException("alignment has no LongSectionSpace");
            DgnModel pm = ls.DgnModel;
            if (pm == null) throw new InvalidOperationException("LongSectionSpace.DgnModel is null");
            int viewIndex = (int)Num(args, "viewIndex", 1.0);   // default View 2 (1) so plan stays in View 1
            if (viewIndex < 0) viewIndex = 0;
            string step = "SetProfileView";
            try
            {
                LongSectionViewManager.Manager.SetProfileView(viewIndex, pm);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("step=" + step + ": " + ex.GetType().Name + ": " + ex.Message);
            }
            return "OK\nkey\tvalue\nalignment\t" + Safe(() => a.Name) + "\nviewIndex\t" + viewIndex +
                   "\nprofileModel\t" + Safe(() => pm.ModelName) + "\nprofileModelId\t" + Safe(() => pm.GetModelId().ToString());
        }

        // N4: Element Annotation on the alignment's profile model (PersistenceAnnotationProcessor).
        // Guard: alignment name must start with WZTC_TEST_ unless force=1. role defaults to "Profile Annotation".
        // dryRun=1 only resolves types/ctor (no Annotate call).
        static string AnnotateProfile(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            string aname = Safe(() => a.Name);
            if (!aname.StartsWith("WZTC_TEST_") && Str(args, "force", "0") != "1")
                throw new InvalidOperationException("refused: annotate_profile only on WZTC_TEST_* unless force=1 (got '" + aname + "')");
            LongSectionSpace ls = a.LongSectionSpace;
            if (ls == null) throw new InvalidOperationException("alignment has no LongSectionSpace");
            DgnModel pm = ls.DgnModel;
            if (pm == null) throw new InvalidOperationException("LongSectionSpace.DgnModel is null");
            // Prefs evidence: Element Annotation last group was E_Stationing; Bulk annotate used Profile Grid.
            // Playbook N4 says "Profile Annotation" (EC display label); Create Drawing uses Profile Grid.
            string role = Str(args, "role", "Profile Grid");
            bool dryRun = Str(args, "dryRun", "0") == "1";
            bool setGroup = Str(args, "setGroup", dryRun ? "0" : "1") == "1";
            string step = "resolve";
            try
            {
                Type procType = null, dbType = null, dbEnumType = null;
                foreach (System.Reflection.Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
                {
                    string nm = asm.GetName().Name ?? "";
                    if (nm.IndexOf("Annotation", StringComparison.OrdinalIgnoreCase) < 0) continue;
                    if (procType == null) procType = asm.GetType("Bentley.CifNET.Annotation.PersistenceAnnotationProcessor");
                    if (dbType == null) dbType = asm.GetType("Bentley.CifNET.Annotation.DrawingBoundary");
                    if (dbEnumType == null) dbEnumType = asm.GetType("Bentley.CifNET.Annotation.DrawingBoundaryType");
                }
                if (procType == null) throw new InvalidOperationException("PersistenceAnnotationProcessor type not loaded");
                step = "SetSelectedAnnotationGroup";
                string groupSet = "skipped";
                if (setGroup && dbType != null && dbEnumType != null)
                {
                    object db = Activator.CreateInstance(dbType);
                    object left = Enum.Parse(dbEnumType, "Left");
                    System.Reflection.MethodInfo set = dbType.GetMethod("SetSelectedAnnotationGroup", new Type[] { typeof(string), dbEnumType });
                    if (set != null) { set.Invoke(db, new object[] { role, left }); groupSet = "Left:" + role; }
                    else groupSet = "no SetSelectedAnnotationGroup method";
                }
                step = "ctor";
                // Processor is bound to the profile model (Element Annotation runs in the profile view).
                object proc = Activator.CreateInstance(procType, new object[] { pm });
                DgnModel activeDm = ActiveDgnModel();
                int beforeProf = CountGraphical(pm), beforeActive = CountGraphical(activeDm);
                if (dryRun)
                {
                    return "OK\nkey\tvalue\nalignment\t" + aname + "\nprofileModelId\t" + Safe(() => pm.GetModelId().ToString()) +
                           "\nrole\t" + role + "\ngroupSet\t" + groupSet +
                           "\nelementsBeforeProfile\t" + beforeProf + "\nelementsBeforeActive\t" + beforeActive +
                           "\ndryRun\t1\nprocType\t" + procType.FullName;
                }
                // target: profile | alignment | model | all  (default all = try every EC + model)
                string target = Str(args, "target", "all");
                System.Reflection.MethodInfo annEc = null;
                foreach (System.Reflection.MethodInfo mi in procType.GetMethods())
                {
                    if (mi.Name != "Annotate") continue;
                    System.Reflection.ParameterInfo[] ps = mi.GetParameters();
                    if (ps.Length == 2 && ps[1].ParameterType == typeof(string) && ps[0].ParameterType.Name.IndexOf("ECInstance") >= 0)
                    { annEc = mi; break; }
                }
                System.Reflection.PropertyInfo ecProp = typeof(ConsensusItem).GetProperty("DgnECInstance");
                var statuses = new StringBuilder();
                Profile ap = a.ActiveProfile;
                if ((target == "profile" || target == "all") && ap != null && annEc != null && ecProp != null)
                {
                    step = "Annotate(profile)";
                    object ec = ecProp.GetValue((ConsensusItem)ap, null);
                    if (ec != null) statuses.Append("profile=" + Convert.ToString(annEc.Invoke(proc, new object[] { ec, role }), System.Globalization.CultureInfo.InvariantCulture) + ";");
                    else statuses.Append("profile=nullEC;");
                }
                if ((target == "alignment" || target == "all") && annEc != null && ecProp != null)
                {
                    step = "Annotate(alignment)";
                    object ec = ecProp.GetValue((ConsensusItem)a, null);
                    if (ec != null) statuses.Append("alignment=" + Convert.ToString(annEc.Invoke(proc, new object[] { ec, role }), System.Globalization.CultureInfo.InvariantCulture) + ";");
                    else statuses.Append("alignment=nullEC;");
                }
                if (target == "model" || target == "all")
                {
                    step = "Annotate(DgnModel)";
                    System.Reflection.MethodInfo ann = procType.GetMethod("Annotate", new Type[] { typeof(DgnModel), typeof(string) });
                    if (ann != null) statuses.Append("model=" + Convert.ToString(ann.Invoke(proc, new object[] { pm, role }), System.Globalization.CultureInfo.InvariantCulture) + ";");
                }
                int afterProf = CountGraphical(pm), afterActive = CountGraphical(activeDm);
                return "OK\nkey\tvalue\nalignment\t" + aname + "\nprofileModelId\t" + Safe(() => pm.GetModelId().ToString()) +
                       "\nrole\t" + role + "\ngroupSet\t" + groupSet + "\ntarget\t" + target +
                       "\nstatuses\t" + statuses.ToString() +
                       "\nelementsBeforeProfile\t" + beforeProf + "\nelementsAfterProfile\t" + afterProf +
                       "\nelementsDeltaProfile\t" + (afterProf - beforeProf) +
                       "\nelementsBeforeActive\t" + beforeActive + "\nelementsAfterActive\t" + afterActive +
                       "\nelementsDeltaActive\t" + (afterActive - beforeActive);
            }
            catch (Exception ex)
            {
                Exception e = ex is System.Reflection.TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
                throw new InvalidOperationException("step=" + step + ": " + e.GetType().Name + ": " + e.Message);
            }
        }

        // READ-ONLY: dump AnnotationManager.GetSupportedAliases keys (group/role name leads)
        static string ListAnnotationAliases()
        {
            Type mgrType = null;
            foreach (System.Reflection.Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                if ((asm.GetName().Name ?? "").IndexOf("Annotation", StringComparison.OrdinalIgnoreCase) < 0) continue;
                mgrType = asm.GetType("Bentley.CifNET.Annotation.AnnotationManager");
                if (mgrType != null) break;
            }
            if (mgrType == null) throw new InvalidOperationException("AnnotationManager type not loaded");
            System.Reflection.PropertyInfo mgrProp = mgrType.GetProperty("Manager", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
            object mgr = mgrProp == null ? null : mgrProp.GetValue(null, null);
            if (mgr == null) throw new InvalidOperationException("AnnotationManager.Manager is null");
            System.Reflection.MethodInfo get = mgrType.GetMethod("GetSupportedAliases");
            object dict = get == null ? null : get.Invoke(mgr, null);
            var sb = new StringBuilder("OK\nkey\tvalue\n");
            if (dict == null) { sb.Append("aliases\tnone\n"); return sb.ToString().TrimEnd('\n'); }
            int n = 0;
            foreach (System.Collections.DictionaryEntry de in (System.Collections.IDictionary)dict)
            {
                sb.Append(Convert.ToString(de.Key, System.Globalization.CultureInfo.InvariantCulture).Replace("\t", " ") + TAB +
                          Convert.ToString(de.Value, System.Globalization.CultureInfo.InvariantCulture).Replace("\t", " ").Replace("\n", " ") + "\n");
                if (++n > 200) break;
            }
            sb.Append("count\t" + n + "\n");
            return sb.ToString().TrimEnd('\n');
        }

        static int CountGraphical(DgnModel dm)
        {
            int n = 0;
            ModelElementsCollection col = dm.GetGraphicElements();
            if (col == null) return 0;
            foreach (Element el in col) n++;
            return n;
        }

        // alignment -> (station, elevation) samples draped on a terrain (tutorial S4: Quick Profile From Surface)
        static string DrapeTerrainName = ""; static double DrapeFrom = 0, DrapeTo = 0;
        static List<double[]> DrapeAlignment(GeometricModel gm, Alignment a, int terrainIdx, double step)
        {
            TerrainSurface ts = null; int i = 0;
            if (DrapeTerrainName.Length > 0)
            {
                foreach (GeometricModel g in ConsensusConnectionExtensions.GetAllGeometricModels(Con()))
                {
                    foreach (TerrainSurface t in g.TerrainSurfaces) { if (ts == null && Safe(() => t.Name) == DrapeTerrainName) ts = t; }
                }
                if (ts == null) throw new InvalidOperationException("terrain '" + DrapeTerrainName + "' not found");
            }
            else
            {
                foreach (TerrainSurface t in gm.TerrainSurfaces) { if (i++ == terrainIdx) { ts = t; break; } }
                if (ts == null) throw new InvalidOperationException("terrain idx " + terrainIdx + " not found");
            }
            var dtm = ts.DTM;
            Bentley.CifNET.LinearGeometry.LinearElement le = a.LinearGeometry;
            double len = le.Length, st0 = a.Stationing.StartStation;
            var res = new List<double[]>();
            double dFrom = DrapeFrom, dTo = DrapeTo > 0 ? Math.Min(DrapeTo, len) : len;
            for (double d = dFrom; ; d += step)
            {
                double dd = Math.Min(d, dTo);
                var lp = le.GetPointAtDistanceOffset(dd, 0.0);
                var dp = dtm.DrapePoint(new DPoint3d(lp.Coordinates.X, lp.Coordinates.Y, 0.0));
                if (dp != null) res.Add(new double[] { st0 + dd, dp.Coordinates.Z, (double)(int)dp.Code });
                if (dd >= dTo) break;
            }
            return res;
        }

        static string DrapePreview(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            DrapeTerrainName = Str(args, "terrain", ""); DrapeFrom = Num(args, "dFrom", 0); DrapeTo = Num(args, "dTo", 0);
            var pts = DrapeAlignment(gm, a, (int)Num(args, "terrainIdx", 0), Num(args, "step", 100));
            var sb = new StringBuilder("OK\nstation\televation\tcode\n");
            foreach (double[] p in pts) sb.Append(R(p[0]) + "\t" + R(p[1]) + "\t" + p[2] + "\n");
            return sb.ToString().TrimEnd('\n');
        }

        // READ-ONLY: every geometric model the connection knows (active file + references)
        static string AllGeoModels()
        {
            var sb = new StringBuilder("OK\nmodel\tfile\talignments\tterrains\tcorridors\tdetail\n");
            foreach (GeometricModel g in ConsensusConnectionExtensions.GetAllGeometricModels(Con()))
            {
                GeometricModel gg = g;
                var det = new StringBuilder();
                try { foreach (TerrainSurface t in gg.TerrainSurfaces) det.Append("T:" + Safe(() => t.Name) + " "); } catch (Exception) { }
                try { foreach (Alignment a in gg.NamedAlignments) det.Append("A:" + Safe(() => a.Name) + "(" + Safe(() => a.LinearGeometry.Length) + ") "); } catch (Exception) { }
                sb.Append(Safe(() => gg.DgnModel.ModelName) + TAB + Safe(() => gg.DgnModel.GetDgnFile().GetFileName()) + TAB + Safe(() => Count(gg.Alignments)) + TAB +
                          Safe(() => Count(gg.TerrainSurfaces)) + TAB + Safe(() => Count(gg.Corridors)) + TAB + det.ToString().Trim() + "\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        // tutorial S4 committed: drape the alignment on a terrain and add the result as a new profile (off-terrain samples dropped)
        static string ProfileFromTerrain(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            DrapeTerrainName = Str(args, "terrain", ""); DrapeFrom = Num(args, "dFrom", 0); DrapeTo = Num(args, "dTo", 0);
            var samples = DrapeAlignment(gm, a, (int)Num(args, "terrainIdx", 0), Num(args, "step", 25));
            var pts = new List<DPoint3d>();
            double lastX = double.MinValue;
            foreach (double[] p in samples)
            {
                if ((int)p[2] != 1) continue;           // only points on the terrain
                if (p[0] <= lastX + 1e-6) continue;     // strictly increasing stations
                pts.Add(new DPoint3d(p[0], p[1], 0.0)); lastX = p[0];
            }
            if (pts.Count < 2) throw new InvalidOperationException("fewer than 2 on-terrain samples");
            int before = Count(a.Profiles);
            Con().StartTransientMode();
            Bentley.CifNET.LinearGeometry.ProfileElement pe = Bentley.CifNET.LinearGeometry.ProfileElement.CreateByStartStationEndStationPassThruPoints(pts[0].X, pts[pts.Count - 1].X, pts.ToArray());
            Profile pr = a.CreateProfileByProfileElement(pe, false, Str(args, "ruled", "1") == "1");
            if (pr == null) throw new InvalidOperationException("CreateProfileByProfileElement returned null");
            Persist(pr);
            return "OK\nkey\tvalue\nalignment\t" + Safe(() => a.Name) + "\nsamples\t" + samples.Count + "\nused\t" + pts.Count + "\nstationFrom\t" + R(pts[0].X) + "\nstationTo\t" + R(pts[pts.Count - 1].X) + "\nprofilesBefore\t" + before + "\nprofilesAfter\t" + Count(a.Profiles);
        }

        static void DumpBoundaries(DgnModel dm, string label, StringBuilder sb)
        {
            try
            {
                foreach (NamedBoundaryGroup g in new NamedBoundaryGroupCollection(dm))
                {
                    NamedBoundaryGroup gg = g;
                    sb.Append(label + TAB + "group=" + Safe(() => gg.Name) + "; count=" + Safe(() => gg.GetCount()) + "\n");
                    foreach (NamedBoundary b in new NamedBoundaryCollection(dm, gg))
                    {
                        NamedBoundary bb = b;
                        sb.Append(label + TAB + "  boundary=" + Safe(() => bb.Name) + "; scale=" + Safe(() => bb.DrawingScale) + "; desc=" + Safe(() => bb.Description) +
                                  "; elemType=" + Safe(() => bb.GraphicalElement.ElementType) + "; level=" + Safe(() => bb.GraphicalElement.LevelId) + "\n");
                        try
                        {
                            ProfileNamedBoundary pnb = ProfileNamedBoundary.CreateFromDgnModel(dm);
                            if (pnb != null) sb.Append(label + TAB + "  profileBoundary=" + Safe(() => pnb.Name) + "; group=" + Safe(() => pnb.GroupName) + "; assocAlignment=" + Safe(() => pnb.GetAssociatedAlignment().Name) + "\n");
                        }
                        catch (Exception) { }
                    }
                }
            }
            catch (Exception ex) { sb.Append(label + TAB + "(boundaries: " + ex.GetType().Name + ": " + ex.Message + ")\n"); }
        }

        // READ-ONLY: public types whose name contains `filter` in the loaded Bentley assemblies, with constructors and methods
        // (used to find the API for things the COM object model cannot enumerate, e.g. saved views).
        static string ProbeTypes(Dictionary<string, object> args)
        {
            string filter = Str(args, "filter", "NamedView");
            var sb = new StringBuilder("OK\ntype\tdetail\n");
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                string an = asm.GetName().Name;
                if (!an.StartsWith("Bentley.DgnPlatformNET") && !an.StartsWith("Bentley.MstnPlatformNET") && !an.StartsWith("ustation")) continue;
                Type[] types;
                try { types = asm.GetTypes(); } catch (System.Reflection.ReflectionTypeLoadException ex) { types = ex.Types; }
                foreach (var t in types)
                {
                    if (t == null || !t.IsPublic || t.FullName == null || t.FullName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    sb.Append(t.FullName + "\t[" + an + "]\n");
                    foreach (var c in t.GetConstructors())
                    {
                        var pts = new List<string>();
                        foreach (var pi in c.GetParameters()) pts.Add(pi.ParameterType.Name + " " + pi.Name);
                        sb.Append("  ctor\t(" + string.Join(", ", pts.ToArray()) + ")\n");
                    }
                    int n = 0;
                    foreach (var m in t.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.DeclaredOnly))
                    {
                        if (n++ >= 30) break;
                        var pts = new List<string>();
                        foreach (var pi in m.GetParameters()) pts.Add(pi.ParameterType.Name + " " + pi.Name);
                        sb.Append("  " + (m.IsStatic ? "static " : "") + m.ReturnType.Name + " " + m.Name + "(" + string.Join(", ", pts.ToArray()) + ")\n");
                    }
                    foreach (var pr in t.GetProperties()) sb.Append("  prop\t" + pr.PropertyType.Name + " " + pr.Name + "\n");
                }
            }
            return sb.ToString().TrimEnd('\n');
        }

        // Saved (named) views of the active file. dry=true (default) lists; with dry=false deletes the ones whose name starts with
        // `prefix` (WZTC_TEST*) and is not in `keep` (comma list of exact names, or a name starting with one of them + " ").
        static string NamedViewsOp(Dictionary<string, object> args)
        {
            DgnFile f = ActiveDgnModel().GetDgnFile();
            string prefix = Str(args, "prefix", "WZTC_TEST");
            bool dry = Str(args, "dry", "true").ToLowerInvariant() != "false";
            var keep = new List<string>();
            foreach (string k in Str(args, "keep", "").Split(',')) if (k.Trim().Length > 0) keep.Add(k.Trim());
            var sb = new StringBuilder("OK\nkey\tvalue\n");
            object coll = null;
            foreach (var pr in typeof(DgnFile).GetProperties()) if (coll == null && pr.PropertyType == typeof(NamedViewCollection)) coll = pr.GetValue(f, null);
            if (coll == null)
                foreach (var mi in typeof(DgnFile).GetMethods())
                    if (coll == null && mi.ReturnType == typeof(NamedViewCollection) && mi.GetParameters().Length == 0) coll = mi.Invoke(f, null);
            if (coll == null)
                foreach (var ci in typeof(NamedViewCollection).GetConstructors(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance))
                {
                    var ps = ci.GetParameters();
                    if (coll == null && ps.Length == 1 && ps[0].ParameterType == typeof(DgnFile)) coll = ci.Invoke(new object[] { f });
                }
            if (coll == null)
            {
                var names = new List<string>();
                foreach (var mi in typeof(DgnFile).GetMethods()) if (mi.Name.IndexOf("View", StringComparison.OrdinalIgnoreCase) >= 0) names.Add(mi.ReturnType.Name + " " + mi.Name);
                throw new InvalidOperationException("no way to get a NamedViewCollection; DgnFile view methods: " + string.Join("; ", names.ToArray()));
            }
            var all = new List<NamedView>();
            foreach (object o in (System.Collections.IEnumerable)coll) all.Add((NamedView)o);
            sb.Append("count\t" + all.Count + "\n");
            int del = 0;
            foreach (NamedView nv in all)
            {
                string nm = Safe(() => nv.Name);
                bool kept = false;
                foreach (string k in keep) if (nm == k || nm.StartsWith(k + " ")) kept = true;
                bool target = nm.StartsWith(prefix) && !kept;
                sb.Append((target ? (dry ? "would delete\t" : "delete\t") : "keep\t") + nm + "\n");
                if (target && !dry) { sb.Append("  status\t" + nv.DeleteFromFile() + "\n"); del++; }
            }
            if (!dry) sb.Append("deletedCount\t" + del + "\n");
            return sb.ToString().TrimEnd('\n');
        }

        // Delete named boundaries (and their groups when emptied) from an alignment's profile model, keeping the names in `keep`.
        // dry=true (default) only lists what would go and which Delete-like methods the API offers. Only WZTC_TEST* names are touched.
        static string DeleteNamedBoundaries(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            DgnModel pm = a.LongSectionSpace.DgnModel;
            var keep = new HashSet<string>();
            foreach (string k in Str(args, "keep", "").Split(',')) if (k.Trim().Length > 0) keep.Add(k.Trim());
            bool dry = Str(args, "dry", "true").ToLowerInvariant() != "false";
            var also = new HashSet<string>();                       // extra exact names allowed besides the WZTC_TEST* prefix
            foreach (string k in Str(args, "also", "").Split(',')) if (k.Trim().Length > 0) also.Add(k.Trim());
            var sb = new StringBuilder("OK\nkey\tvalue\n");
            var del = new List<NamedBoundary>();
            var groups = new List<NamedBoundaryGroup>();
            foreach (NamedBoundaryGroup g in new NamedBoundaryGroupCollection(pm))
            {
                groups.Add(g);
                foreach (NamedBoundary b in new NamedBoundaryCollection(pm, g))
                {
                    string bn = Safe(() => b.Name);
                    if ((bn.StartsWith("WZTC_TEST") || also.Contains(bn)) && !keep.Contains(bn)) del.Add(b);
                }
            }
            sb.Append("candidates\t" + del.Count + "\n");
            foreach (var mi in typeof(NamedBoundary).GetMethods())
                if (mi.Name.IndexOf("Delete", StringComparison.OrdinalIgnoreCase) >= 0 || mi.Name.IndexOf("Remove", StringComparison.OrdinalIgnoreCase) >= 0)
                    sb.Append("NamedBoundary.method\t" + mi.Name + "(" + mi.GetParameters().Length + ")\n");
            foreach (var mi in typeof(NamedBoundaryGroup).GetMethods())
                if (mi.Name.IndexOf("Delete", StringComparison.OrdinalIgnoreCase) >= 0 || mi.Name.IndexOf("Remove", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var pts = new List<string>();
                    foreach (var pi in mi.GetParameters()) pts.Add(pi.ParameterType.FullName);
                    sb.Append("NamedBoundaryGroup.method\t" + mi.Name + "(" + string.Join(", ", pts.ToArray()) + ")\n");
                }
            if (dry) { foreach (var b in del) sb.Append("would delete\t" + Safe(() => b.Name) + "\n"); return sb.ToString().TrimEnd('\n'); }
            int n = 0;
            foreach (NamedBoundary b in del)
            {
                string bn = Safe(() => b.Name);
                object r = null; string how = "";
                var m0 = typeof(NamedBoundary).GetMethod("Delete", Type.EmptyTypes) ?? typeof(NamedBoundary).GetMethod("DeleteFromModel", Type.EmptyTypes);
                if (m0 != null) { r = m0.Invoke(b, null); how = m0.Name; }
                else
                {
                    var ge = b.GraphicalElement;
                    var m1 = ge.GetType().GetMethod("DeleteFromModel", Type.EmptyTypes);
                    if (m1 == null) throw new InvalidOperationException("no delete method for boundary " + bn);
                    r = m1.Invoke(ge, null); how = "GraphicalElement.DeleteFromModel";
                }
                sb.Append("deleted\t" + bn + " via " + how + " -> " + r + "\n"); n++;
            }
            sb.Append("deletedCount\t" + n + "\n");
            // groups left empty by the deletions (never a group that still holds a kept boundary)
            var empty = new List<NamedBoundaryGroup>();
            foreach (NamedBoundaryGroup g in new NamedBoundaryGroupCollection(pm))
            {
                string gn = Safe(() => g.Name);
                if (gn.StartsWith("WZTC_TEST") && Safe(() => g.GetCount()) == "0") empty.Add(g);
            }
            foreach (NamedBoundaryGroup g in empty)
            {
                string gn = Safe(() => g.Name);
                System.Reflection.MethodInfo gm1 = null;
                foreach (var mi in typeof(NamedBoundaryGroup).GetMethods())
                    if (mi.Name == "Delete" && mi.GetParameters().Length == 1) gm1 = mi;
                if (gm1 == null) { sb.Append("group not deleted (no Delete with one argument)\t" + gn + "\n"); continue; }
                Type pt = gm1.GetParameters()[0].ParameterType;
                object arg = null;
                if (pt == typeof(DgnModel)) arg = pm;
                else if (pt == typeof(DgnFile)) arg = pm.GetDgnFile();
                else if (pt == typeof(bool)) arg = true;
                else if (pt.IsInstanceOfType(g)) arg = g;
                if (arg == null) { sb.Append("group not deleted (unsupported argument " + pt.FullName + ")\t" + gn + "\n"); continue; }
                sb.Append("group deleted\t" + gn + " -> " + gm1.Invoke(g, new object[] { arg }) + "\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        // READ-ONLY: named boundaries (platform + Civil Profile view) in the alignment's profile model and every model of the file
        static string ProfileBoundaries(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            var sb = new StringBuilder("OK\nmodel\tinfo\n");
            DgnModel pm = a.LongSectionSpace.DgnModel;
            DumpBoundaries(pm, "profileModel(" + Safe(() => a.Name) + ")", sb);
            DgnFile f = ActiveDgnModel().GetDgnFile();
            foreach (object it in f.GetModelIndexCollection())
            {
                string mn = Convert.ToString(it.GetType().GetProperty("Name").GetValue(it, null));
                try { DumpBoundaries(LoadByName(f, mn), mn, sb); } catch (Exception ex) { sb.Append(mn + TAB + "(" + ex.GetType().Name + ")\n"); }
            }
            return sb.ToString().TrimEnd('\n');
        }

        // READ-ONLY: Civil Profile boundary data of an alignment's profile model (what the Place Named Boundary dialog stored)
        static string ProfileBoundaryData(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            DgnModel pm = a.LongSectionSpace.DgnModel;
            var sb = new StringBuilder("OK\nkey\tvalue\n");
            string step = "CreateFromDgnModel";
            try
            {
                ProfileNamedBoundary pnb = ProfileNamedBoundary.CreateFromDgnModel(pm);
                DgnFile pf = ActiveDgnModel().GetDgnFile();
                // optional: read the boundary held by one named model (the Drawing model Create Drawing made); no fallback to another boundary
                string wantModel = args.ContainsKey("model") ? Convert.ToString(args["model"]) : "";
                if (wantModel.Length > 0)
                {
                    pnb = ProfileNamedBoundary.CreateFromDgnModel(LoadByName(pf, wantModel));
                    sb.Append("viaModel[" + wantModel + "]\t" + (pnb == null ? "null" : "present") + "\n");
                }
                else
                foreach (object it in pf.GetModelIndexCollection())
                {
                    string mn = Convert.ToString(it.GetType().GetProperty("Name").GetValue(it, null));
                    try { ProfileNamedBoundary q = ProfileNamedBoundary.CreateFromDgnModel(LoadByName(pf, mn)); sb.Append("viaModel[" + mn + "]\t" + (q == null ? "null" : "present") + "\n"); if (pnb == null && q != null) pnb = q; }
                    catch (Exception ex) { sb.Append("viaModel[" + mn + "]\t" + ex.GetType().Name + ": " + ex.Message + "\n"); }
                }
                sb.Append("profileBoundary\t" + (pnb == null ? "null" : "present") + "\n");
                if (pnb != null)
                {
                    sb.Append("name\t" + Safe(() => pnb.Name) + "\ngroup\t" + Safe(() => pnb.GroupName) + "\nassocAlignment\t" + Safe(() => pnb.GetAssociatedAlignment().Name) + "\n");
                    step = "GetProfileData";
                    double sx = 0, sy = 0, sd = 0, ed = 0, len = 0, ph = 0, tc = 0, bc = 0, ve = 0, vs = 0, hs = 0, fx = 0, fy = 0; bool ut = false, ub = false, dt = false, dv = false; int shift = 0, idx = 0, num = 0; string pid = "";
                    pnb.GetProfileData(ref sx, ref sy, ref sd, ref ed, ref len, ref ph, ref tc, ref bc, ref ut, ref ub, ref dt, ref dv, ref shift, ref ve, ref vs, ref hs, ref pid, ref idx, ref num, ref fx, ref fy);
                    sb.Append("startOffsetX\t" + R(sx) + "\nstartOffsetY\t" + R(sy) + "\nstartDistance\t" + R(sd) + "\nstopDistance\t" + R(ed) + "\nlength\t" + R(len) +
                              "\nprofileHeight\t" + R(ph) + "\ntopClearance\t" + R(tc) + "\nbotClearance\t" + R(bc) + "\nuseTopClearance\t" + ut + "\nuseBotClearance\t" + ub +
                              "\ndisplayTerrain\t" + dt + "\ndisplayVertical\t" + dv + "\nshiftLocation\t" + shift + "\nverticalExaggeration\t" + R(ve) +
                              "\nverticalAxisSpacing\t" + R(vs) + "\nhorizontalAxisSpacing\t" + R(hs) + "\nsegmentParentID\t" + pid + "\nsegmentIndex\t" + idx + "\nnumSegments\t" + num + "\n");
                }
            }
            catch (Exception ex) { sb.Append("error\tstep=" + step + ": " + ex.GetType().Name + ": " + ex.Message + "\n"); }
            return sb.ToString().TrimEnd('\n');
        }

        // READ-ONLY: which ProfileElement factory gives straight segments between the points? (nothing is attached or persisted)
        static string ProbeProfile2()
        {
            var pts = new DPoint3d[] { new DPoint3d(0, 0, 0), new DPoint3d(100, 10, 0), new DPoint3d(200, 0, 0), new DPoint3d(300, 30, 0) };
            var sb = new StringBuilder("OK\nfactory\ty@50,150,250 (linear would be 5,5,15)\n");
            string[] names = { "CreateByStartStationEndStationPassThruPoints", "Create1", "Create2" };
            for (int v = 0; v < 3; v++)
            {
                string r;
                try
                {
                    Bentley.CifNET.LinearGeometry.ProfileElement pe = v == 0 ? Bentley.CifNET.LinearGeometry.ProfileElement.CreateByStartStationEndStationPassThruPoints(0, 300, pts)
                        : v == 1 ? Bentley.CifNET.LinearGeometry.ProfileElement.Create1(0, 300, pts) : Bentley.CifNET.LinearGeometry.ProfileElement.Create2(0, 300, pts);
                    Bentley.CifNET.LinearGeometry.ProfileElement q = pe;
                    r = Safe(() => q.GetYAtX(50)) + ", " + Safe(() => q.GetYAtX(150)) + ", " + Safe(() => q.GetYAtX(250)) + "  len=" + Safe(() => q.Length);
                }
                catch (Exception ex) { r = ex.GetType().Name + ": " + ex.Message; }
                sb.Append(names[v] + "\t" + r + "\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        // delete ONE profile of an alignment. Guard: expectX0 must equal the profile's first x (station) so the wrong profile can't go.
        static string DeleteProfile(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            int want = (int)Num(args, "profileIdx", -1);
            double expect = Num(args, "expectX0", double.NaN);
            if (want < 1) throw new InvalidOperationException("profileIdx must be >= 1 (0 is the stub profile)");
            if (double.IsNaN(expect)) throw new InvalidOperationException("expectX0 required (first station of the profile to delete)");
            Profile target = null; int i = 0;
            foreach (Profile p in a.Profiles) { if (i++ == want) { target = p; break; } }
            if (target == null) throw new InvalidOperationException("profileIdx " + want + " not found");
            double x0 = target.ProfileGeometry.ProjectedDistRange.Item1;
            if (Math.Abs(x0 - expect) > 0.01) throw new InvalidOperationException("refused: profile " + want + " starts at " + R(x0) + ", expected " + R(expect));
            int before = Count(a.Profiles);
            Con().Delete((ConsensusItem)target);
            return "OK\nkey\tvalue\nalignment\t" + Safe(() => a.Name) + "\nprofilesBefore\t" + before + "\nprofilesAfter\t" + Count(a.Profiles);
        }

        // flush any pending transient edits so the connection returns to normal mode
        static string FlushTransients()
        {
            Con().PersistTransients();
            return "OK\nkey\tvalue\nflushed\ttrue";
        }

        // delete the profiles with index >= fromIdx whose range lies inside [winX0, winX1] and is at most maxLen long (cleanup of segment profiles)
        static string DeleteProfilesWindow(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            int fromIdx = (int)Num(args, "fromIdx", 3);
            if (fromIdx < 3 && !Safe(() => a.Name).StartsWith("WZTC_TEST_")) throw new InvalidOperationException("fromIdx must be >= 3 (profiles 0-2 may be the engineer's) unless the alignment is a WZTC_TEST_ object");
            if (fromIdx < 1) throw new InvalidOperationException("profile 0 is the stub");
            double w0 = Num(args, "winX0", double.NaN), w1 = Num(args, "winX1", double.NaN), maxLen = Num(args, "maxLen", 30);
            if (double.IsNaN(w0) || double.IsNaN(w1)) throw new InvalidOperationException("winX0 and winX1 required");
            var targets = new List<Profile>(); var skipped = new List<int>(); int i = 0;
            foreach (Profile p in a.Profiles)
            {
                if (i >= fromIdx)
                {
                    var r = p.ProfileGeometry.ProjectedDistRange;
                    if (r.Item1 >= w0 - 0.01 && r.Item2 <= w1 + 0.01 && (r.Item2 - r.Item1) <= maxLen) targets.Add(p); else skipped.Add(i);
                }
                i++;
            }
            int before = Count(a.Profiles), deleted = 0;
            foreach (Profile p in targets) { Con().Delete((ConsensusItem)p); deleted++; }
            return "OK\nkey\tvalue\nalignment\t" + Safe(() => a.Name) + "\nprofilesBefore\t" + before + "\ndeleted\t" + deleted + "\nskippedIdx\t" + string.Join(",", skipped.ConvertAll(x => x.ToString()).ToArray()) + "\nprofilesAfter\t" + Count(a.Profiles);
        }

        // READ-ONLY: reflect a type (fields, properties, ctors, enum names) from inside the ORD process
        static string Reflect(Dictionary<string, object> args)
        {
            string typeName = Str(args, "type", "");
            Type t = typeof(Alignment).Assembly.GetType(typeName);
            if (t == null)
            {
                foreach (System.Reflection.Assembly a in AppDomain.CurrentDomain.GetAssemblies())
                {
                    try { t = a.GetType(typeName); } catch (Exception) { t = null; }
                    if (t != null) break;
                    if (!string.IsNullOrEmpty(typeName) && typeName.IndexOf('.') < 0)
                    {
                        try
                        {
                            foreach (Type x in a.GetTypes())
                            {
                                if (x.Name == typeName) { t = x; break; }
                            }
                        }
                        catch (Exception) { }
                        if (t != null) break;
                    }
                }
            }
            if (t == null) throw new InvalidOperationException("type not found");
            var sb = new StringBuilder("OK\nkind\tdetail\n");
            var fl = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.DeclaredOnly;
            if (t.IsEnum) { foreach (string n in Enum.GetNames(t)) sb.Append("enum\t" + n + "=" + Convert.ToInt64(Enum.Parse(t, n)) + "\n"); }
            sb.Append("assembly\t" + t.Assembly.GetName().Name + "\n");
            sb.Append("fullname\t" + t.FullName + "\n");
            foreach (System.Reflection.FieldInfo f in t.GetFields(fl)) sb.Append("field\t" + f.FieldType.Name + " " + f.Name + "\n");
            foreach (System.Reflection.PropertyInfo p in t.GetProperties(fl)) sb.Append("prop\t" + p.PropertyType.Name + " " + p.Name + "\n");
            foreach (System.Reflection.ConstructorInfo c in t.GetConstructors(fl)) { var ps = new List<string>(); foreach (System.Reflection.ParameterInfo q in c.GetParameters()) ps.Add(q.ParameterType.Name + " " + q.Name); sb.Append("ctor\t(" + string.Join(", ", ps.ToArray()) + ")\n"); }
            foreach (System.Reflection.MethodInfo m in t.GetMethods(fl))
            {
                if (m.IsSpecialName) continue;
                var ps = new List<string>();
                foreach (System.Reflection.ParameterInfo q in m.GetParameters()) ps.Add(q.ParameterType.Name + " " + q.Name);
                sb.Append("method\t" + m.ReturnType.Name + " " + m.Name + "(" + string.Join(", ", ps.ToArray()) + ")\n");
            }
            return sb.ToString().TrimEnd('\n');
        }

        // ONE profile from vertical PIs with zero-length curves (straight segments): points "station,elev;..."
        static string CreateProfileVPI(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            var pts = new List<DPoint3d>(); var infos = new List<VerticalCurveInfo>();
            foreach (string part in Str(args, "points", "").Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] f = part.Split(',');
                pts.Add(new DPoint3d(double.Parse(f[0], System.Globalization.CultureInfo.InvariantCulture), double.Parse(f[1], System.Globalization.CultureInfo.InvariantCulture), 0.0));
                VerticalCurveInfo vi = new VerticalCurveInfo();
                vi.curveType = VerticalCurveType.Parabola; vi.curveParam = 0; vi.curveLength = 0; vi.slopeRate = 0;
                infos.Add(vi);
            }
            if (pts.Count < 2) throw new InvalidOperationException("need >= 2 points");
            int before = Count(a.Profiles);
            Con().StartTransientMode();
            Profile pr = a.CreateProfileByVerticalPIs(infos, pts, false);
            if (pr == null) throw new InvalidOperationException("CreateProfileByVerticalPIs returned null");
            Persist(pr);
            return "OK\nkey\tvalue\nalignment\t" + Safe(() => a.Name) + "\npoints\t" + pts.Count + "\nprofilesBefore\t" + before + "\nprofilesAfter\t" + Count(a.Profiles);
        }

        // playbook N3 as code: PVI table "station,elevation,curveLength;..." (metres) -> ONE profile with parabolic vertical curves, set ACTIVE.
        // mode: "all" = one VerticalCurveInfo per PVI, "interior" = one per interior PVI. Returns the new profile's index.
        static string CreateProfileFromPVIs(Dictionary<string, object> args)
        {
            GeometricModel gm = Model();
            Alignment a = Pick(gm, args);
            var pts = new List<DPoint3d>(); var lens = new List<double>();
            foreach (string part in Str(args, "pvis", "").Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] f = part.Split(',');
                if (f.Length != 3) throw new InvalidOperationException("each PVI is 'station,elevation,curveLength'");
                pts.Add(new DPoint3d(double.Parse(f[0], System.Globalization.CultureInfo.InvariantCulture), double.Parse(f[1], System.Globalization.CultureInfo.InvariantCulture), 0.0));
                lens.Add(double.Parse(f[2], System.Globalization.CultureInfo.InvariantCulture));
            }
            if (pts.Count < 2) throw new InvalidOperationException("need >= 2 PVIs");
            string mode = Str(args, "mode", "all");
            var infos = new List<VerticalCurveInfo>();
            for (int i = 0; i < pts.Count; i++)
            {
                if (mode == "interior" && (i == 0 || i == pts.Count - 1)) continue;
                double g1 = i > 0 ? (pts[i].Y - pts[i - 1].Y) / (pts[i].X - pts[i - 1].X) : 0, g2 = i < pts.Count - 1 ? (pts[i + 1].Y - pts[i].Y) / (pts[i + 1].X - pts[i].X) : 0;
                double A = Math.Abs(g2 - g1) * 100.0;        // algebraic difference in percent
                VerticalCurveInfo vi = new VerticalCurveInfo();
                vi.curveType = VerticalCurveType.Parabola; vi.curveLength = lens[i];
                vi.curveParam = (lens[i] > 0 && A > 0) ? lens[i] / A : 0;    // K value (length per percent grade change)
                vi.slopeRate = (mode == "slope0") ? 0 : g2;      // grade leaving this PVI
                infos.Add(vi);
            }
            int before = Count(a.Profiles);
            Con().StartTransientMode();
            Profile pr = a.CreateProfileByVerticalPIs(infos, pts, false);
            if (pr == null) throw new InvalidOperationException("CreateProfileByVerticalPIs returned null");
            Persist(pr);
            int after = Count(a.Profiles);
            // the new profile = the one whose range starts at the first PVI and ends at the last (searched from the end)
            int idx = -1, i2 = 0;
            foreach (Profile q in a.Profiles)
            {
                var r = q.ProfileGeometry.ProjectedDistRange;
                if (Math.Abs(r.Item1 - pts[0].X) < 0.01 && Math.Abs(r.Item2 - pts[pts.Count - 1].X) < 0.01) idx = i2;
                i2++;
            }
            string act = "";
            if (idx >= 0 && Str(args, "active", "1") == "1")
            {
                int j = 0;
                foreach (Profile q in a.Profiles) { if (j++ == idx) { act = a.SetActiveProfile(q).ToString(); break; } }
            }
            return "OK\nkey\tvalue\nalignment\t" + Safe(() => a.Name) + "\nprofileIdx\t" + idx + "\nprofilesBefore\t" + before + "\nprofilesAfter\t" + after +
                   "\nsetActive\t" + act + "\nactiveProfile\t" + Safe(() => a.ActiveProfile == null ? "" : a.ActiveProfile.Name) + "\ninfos\t" + infos.Count + "\nmode\t" + mode;
        }
}
}
