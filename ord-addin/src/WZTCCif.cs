// WZTCCif LOADER: the small, stable COM-visible class VBA creates (CreateObject("WZTC.CifBridge")) inside OpenRoadsDesigner.exe.
// It contains NO OpenRoads logic. It finds the newest implementation DLL (bin\impl\WZTCCifImpl.<stamp>.dll), loads it from bytes
// (so the file is never locked and each build gets its own assembly), and forwards Execute(cmd, json) to WZTC.CifImpl.Run.
// => C# changes are picked up on the next call without restarting OpenRoads or touching the registry.
//
// Reply format of Execute: line 1 = "OK<TAB>k=v..." or "ERROR<TAB>note=..." ; optional TSV table after it (header first).
// C# 5 only (system csc.exe).
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace WZTC
{
    [ComVisible(true)]
    [Guid("8D3A5C1E-6F42-4B7A-9E10-2C7B5D9A41F3")]
    [ProgId("WZTC.CifBridge")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class CifBridge
    {
        static readonly string[] Dirs = new string[] {
            @"C:\Program Files\Bentley\OpenRoads Designer 2023.00\OpenRoadsDesigner\Cif",
            @"C:\Program Files\Bentley\OpenRoads Designer 2023.00\OpenRoadsDesigner",
            @"C:\Program Files\Bentley\OpenRoads Designer 2023.00\OpenRoadsDesigner\Assemblies" };
        static readonly HashSet<string> busy = new HashSet<string>();
        static string implPath;
        static MethodInfo implRun;
        static string loaderDir;

        static CifBridge()
        {
            loaderDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            AppDomain.CurrentDomain.AssemblyResolve += delegate(object s, ResolveEventArgs e)
            {
                string n = e.Name.Split(',')[0];
                if (!n.StartsWith("Bentley.")) return null;
                if (!busy.Add(n)) return null;
                try
                {
                    foreach (string d in Dirs)
                    {
                        string p = Path.Combine(d, n + ".dll");
                        if (File.Exists(p)) return Assembly.LoadFrom(p);
                    }
                    return null;
                }
                finally { busy.Remove(n); }
            };
        }

        public string Ping()
        {
            return Execute("ping", "{}");
        }

        public string Execute(string cmd, string json)
        {
            try
            {
                MethodInfo run = Impl();
                string head = "loader=" + Assembly.GetExecutingAssembly().GetName().Version + "\timpl=" + Path.GetFileName(implPath);
                string res = (string)run.Invoke(null, new object[] { cmd, json ?? "{}", head });
                return res;
            }
            catch (TargetInvocationException tie)
            {
                return "ERROR\tnote=" + Flat(tie.InnerException ?? tie);
            }
            catch (Exception ex)
            {
                return "ERROR\tnote=" + Flat(ex);
            }
        }

        // newest bin\impl\WZTCCifImpl.*.dll by name (names carry a sortable timestamp); reload only when it changes
        static MethodInfo Impl()
        {
            string dir = Path.Combine(loaderDir, "impl");
            if (!Directory.Exists(dir)) throw new FileNotFoundException("implementation folder missing: " + dir);
            string best = null;
            foreach (string f in Directory.GetFiles(dir, "WZTCCifImpl.*.dll"))
                if (best == null || string.CompareOrdinal(Path.GetFileName(f), Path.GetFileName(best)) > 0) best = f;
            if (best == null) throw new FileNotFoundException("no WZTCCifImpl.*.dll in " + dir + " (run the build)");
            if (best != implPath || implRun == null)
            {
                Assembly asm = Assembly.Load(File.ReadAllBytes(best));
                Type t = asm.GetType("WZTC.CifImpl", true);
                implRun = t.GetMethod("Run", BindingFlags.Public | BindingFlags.Static);
                if (implRun == null) throw new MissingMethodException("WZTC.CifImpl.Run not found in " + best);
                implPath = best;
            }
            return implRun;
        }

        static string Flat(Exception ex)
        {
            var sb = new StringBuilder();
            for (Exception e = ex; e != null; e = e.InnerException)
            {
                if (sb.Length > 0) sb.Append(" <- ");
                sb.Append(e.GetType().Name + ": " + e.Message);
            }
            return sb.ToString().Replace("\t", " ").Replace("\r", " ").Replace("\n", " ");
        }
    }
}
