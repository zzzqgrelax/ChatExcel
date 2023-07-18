using ChatExcel.Shared.Consts;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ChatExcel.Addin.Utilties
{
    internal class ProcessHelper
    {
        internal static void EnsureProcess()
        {
            try
            {
                var processExisted = Process.GetProcesses()
                     .Any(pr => pr.ProcessName.ToLower().Equals(ProcessConst.ChatExcelProcess.ToLower()));
                if (!processExisted)
                    StartProcess();
            }
            catch (Exception)
            { }

            try
            {
                if (ExcelDnaUtil.IsET && !NamedPipeServerHelper.ServerStarted)
                    NamedPipeServerHelper.InitServer();
            }
            catch { }
        }

        private static void StartProcess()
        {
            var xllPath = ExcelDnaUtil.XllPath;
            var xlldir = Path.GetDirectoryName(xllPath);
            var pro = xlldir + $@"\{ProcessConst.ChatExcelProcess}.exe";
            var startInfo = new ProcessStartInfo
            {
                UseShellExecute = false,
                FileName = pro
            };
            var process = Process.Start(startInfo);
            while (!process.WaitForInputIdle())
            {
                Task.Delay(500).Wait();
            }
        }
    }
}
