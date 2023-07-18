using ChatExcel.Addin.Utilties;
using ChatExcel.Shared.Consts;
using ExcelDna.ComInterop;
using ExcelDna.Integration;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Application = NetOffice.ExcelApi.Application;

namespace ChatExcel.Addin.AddIn
{
    [ComVisible(true)]
    public class AddIn : IExcelAddIn
    {
        internal static Application ExcelApp;
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
            ExcelApp.Dispose();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void AutoOpen()
        {
            HandleUnexpectedExceptions();
            LoadAssemblyDlls();
            InitEvent();
            NamedPipeServerHelper.InitServer();
#if DEBUG
            Log.LogFileDir = AppSettings.logDirectory;
#endif
            ComServer.DllRegisterServer();
            StatusBarMsgHelper.InitStatusBarMsg();
        }

        void HandleUnexpectedExceptions()
        {
            try
            {
                ExcelIntegration.RegisterUnhandledExceptionHandler(obj =>
                {
                    if (obj != null && obj is Exception)
                    {
                        var ex = obj as Exception;
                        var message = string.Format("#EXCEPTION: {0}", ex.Message);
                        return message;
                    }
                    return string.Format("#UNEXPECTED_EXCEPTION: " + obj.ToString());
                });
            }
            catch (Exception) { }
        }

        void LoadAssemblyDlls()
        {
            try
            {
                //https://stackoverflow.com/questions/33109558/revit-api-possible-newtonsoft-json-conflict/39541168#39541168
                string dir = ExcelDnaUtil.XllPathInfo.DirectoryName;
                string dllPath = Path.Combine(dir, "Newtonsoft.Json.dll");
                Assembly.LoadFrom(dllPath);
            }
            catch (Exception) { }
        }

        internal void InitEvent()
        {
            if (ExcelApp == null)
                ExcelApp = new Application(null, ExcelDnaUtil.Application);
            ExcelApp.WorkbookOpenEvent += ExcelApp_WorkbookOpenEvent;
        }

        private void ExcelApp_WorkbookOpenEvent(NetOffice.ExcelApi.Workbook wb)
        {
            try
            {
                XLApp.ScreenUpdate(true);
            }
            catch (Exception) { }

            var processExisted = Process.GetProcesses()
                   .Any(pr => pr.ProcessName.ToLower().Equals(ProcessConst.ChatExcelProcess.ToLower()));
            if (!processExisted)
                return;

            InvokeUtil.QueueAsMacro(() =>
            {
                using (var app = new Application(null, ExcelDnaUtil.Application))
                {
                    //do some init work
                }
            });
        }
    }
}
