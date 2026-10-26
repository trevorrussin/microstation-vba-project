// Implementation entry point, loaded by the WZTCCif loader. Run(cmd, json, head) -> reply (see WZTCCif.cs for the format).
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Web.Script.Serialization;

namespace WZTC
{
    public static class CifImpl
    {
        public static string Run(string cmd, string json, string head)
        {
            try
            {
                Dictionary<string, object> args = new Dictionary<string, object>();
                if (!string.IsNullOrEmpty(json))
                    args = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(json);
                if (cmd == "ping")
                {
                    var sb = new StringBuilder("OK");
                    sb.Append("\tprocess=" + Process.GetCurrentProcess().ProcessName);
                    sb.Append("\tclr=" + Environment.Version);
                    sb.Append("\t" + head);
                    try { sb.Append("\t" + CifOps.Describe()); }
                    catch (Exception ex) { sb.Append("\tcif=UNAVAILABLE: " + Flat(ex)); }
                    return sb.ToString();
                }
                return CifOps.Run(cmd, args);
            }
            catch (Exception ex)
            {
                return "ERROR\tnote=" + Flat(ex);
            }
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
