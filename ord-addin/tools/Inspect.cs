// Offline API explorer for the OpenRoads / MicroStation .NET assemblies (runs as its own process, never inside ORD).
//   Inspect.exe types  <assembly.dll> <regex>            list public types whose full name matches
//   Inspect.exe members <assembly.dll> <full type name> [regex]   list public members (methods with parameters)
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

static class Inspect
{
    static readonly string[] Dirs = new[] {
        @"C:\Program Files\Bentley\OpenRoads Designer 2023.00\OpenRoadsDesigner",
        @"C:\Program Files\Bentley\OpenRoads Designer 2023.00\OpenRoadsDesigner\Assemblies",
        @"C:\Program Files\Bentley\OpenRoads Designer 2023.00\OpenRoadsDesigner\Cif",
        @"C:\Program Files\Bentley\OpenRoads Designer 2023.00\OpenRoadsDesigner\MdlSys\Asneeded",
        @"C:\Program Files\Bentley\OpenRoads Designer 2023.00\OpenRoadsDesigner\MdlSys\Required" };
    static readonly HashSet<string> busy = new HashSet<string>();

    static Assembly Resolve(object s, ResolveEventArgs e)
    {
        string n = e.Name.Split(',')[0];
        if (!busy.Add(n)) return null;            // no recursion
        try
        {
            foreach (var d in Dirs)
            {
                var p = Path.Combine(d, n + ".dll");
                if (File.Exists(p)) return Assembly.LoadFrom(p);
            }
            return null;
        }
        finally { busy.Remove(n); }
    }

    static Type[] Types(Assembly a)
    {
        try { return a.GetTypes(); }
        catch (ReflectionTypeLoadException ex) { return ex.Types.Where(t => t != null).ToArray(); }
    }

    static string Sig(MethodBase m)
    {
        string ps;
        try { ps = string.Join(", ", m.GetParameters().Select(p => p.ParameterType.Name + " " + p.Name)); }
        catch (Exception) { ps = "?"; }
        var mi = m as MethodInfo;
        string ret = mi != null ? SafeName(() => mi.ReturnType.Name) + " " : "";
        return (m.IsStatic ? "static " : "") + ret + m.Name + "(" + ps + ")";
    }

    static string SafeName(Func<string> f) { try { return f(); } catch (Exception) { return "?"; } }

    static int Main(string[] a)
    {
        // mixed-mode (C++/CLI) Bentley assemblies need their native DLLs, which sit next to ORD itself
        Environment.SetEnvironmentVariable("PATH", string.Join(";", Dirs) + ";" + Environment.GetEnvironmentVariable("PATH"));
        AppDomain.CurrentDomain.AssemblyResolve += Resolve;
        if (a.Length < 3) { Console.WriteLine("usage: Inspect types|members <dll> <regex|type> [regex]"); return 2; }
        var asm = Assembly.LoadFrom(a[1]);
        if (a[0] == "types")
        {
            var rx = new Regex(a[2], RegexOptions.IgnoreCase);
            foreach (var t in Types(asm).Where(t => t.IsPublic && rx.IsMatch(t.FullName)).OrderBy(t => t.FullName))
                Console.WriteLine(t.FullName + (t.BaseType != null ? "  : " + SafeName(() => t.BaseType.Name) : ""));
            return 0;
        }
        var type = Types(asm).FirstOrDefault(t => t.FullName == a[2]);
        if (type == null) { Console.WriteLine("type not found: " + a[2]); return 1; }
        var mrx = new Regex(a.Length > 3 ? a[3] : ".", RegexOptions.IgnoreCase);
        const BindingFlags F = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;
        Console.WriteLine("== " + type.FullName + "  base: " + SafeName(() => type.BaseType.FullName));
        foreach (var c in type.GetConstructors())
            Console.WriteLine("  ctor " + Sig(c));
        foreach (var m in type.GetMethods(F).Where(m => !m.IsSpecialName && mrx.IsMatch(m.Name)).OrderBy(m => m.Name))
            Console.WriteLine("  " + Sig(m));
        foreach (var p in type.GetProperties(F).Where(p => mrx.IsMatch(p.Name)).OrderBy(p => p.Name))
            Console.WriteLine("  prop " + SafeName(() => p.PropertyType.Name) + " " + p.Name + (p.CanWrite ? " {get;set;}" : " {get;}"));
        return 0;
    }
}
