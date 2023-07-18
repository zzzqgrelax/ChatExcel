using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChatExcel.Addin.AddIn
{
    static class AppSettings
    {
        internal const string AddinGUID = "32DFF690-5ADA-45F5-A829-F7CEDDED267D";
        internal const string AddinProgID = "ChatExcel.Connect";
        internal const string RibbonGUID = "AD2B0BE5-5A72-4016-934E-A39F1A511D44";
        internal const string RibbonProgID = "ChatExcel.Ribbon";
        internal const string ComRTDGUID = "C70B9BAB-4C5A-431B-8329-DC39AC70D044";

        internal static string logDirectory { get; set; }
        static AppSettings()
        {
            var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ChatExcel", "Excel", "Logs");
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
            logDirectory = directory;
        }
    }
}
