using ChatExcel.Addin.AddIn;
using ChatExcel.Addin.Utilties;
using ChatExcel.Shared.Consts;
using ExcelDna.Integration.CustomUI;
using NamePipeLib;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using System;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ChatExcel.Addin.Ribbon
{
    [ComVisible(true)]
    [Guid(AppSettings.RibbonGUID), ProgId(AppSettings.RibbonProgID)]
    public class RibbonController : ExcelRibbon
    {
        internal static IRibbonUI CustomRibbon;

        public static bool LoggedIn { get; set; } = true;

        public void RibbonLoaded(IRibbonUI ribbon)
        {
            CustomRibbon = ribbon;
        }

        public override string GetCustomUI(string RibbonID)
        {
            return Properties.Resources.CustomUI;
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "LoginButton":
                    if (!LoggedIn)
                        return Properties.Resources.logout;
                    else
                    {
                        return Properties.Resources.login;
                    }
                case "ChatExcelButton":
                    {
                        return Properties.Resources.codeGenerator;
                    }
                case "NavigationMenu":
                    {
                        return Properties.Resources.navigation;
                    }
                case "UpdateButton":
                    {
                        return Properties.Resources.update;
                    }
                case "SaveCopyButton":
                    {
                        return Properties.Resources.saveCopy;
                    }
                default:
                    break;
            }
            if (control.Id.StartsWith("Btn"))
                return Properties.Resources.worksheet;
            if (control.Id.StartsWith("Menu"))
                return Properties.Resources.workBook;
            return null;
        }

        public string GetLabel(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "LoginButton":
                    return LoggedIn ? "已登录\r\n" : "未登录\r\n";
                default:
                    return control.Id;
            }
        }

        public string GetNavigations(IRibbonControl control)
        {
            var app = Application.GetActiveInstance();
            if (app == null)
                return @"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui""></menu>";
            var workBooks = app.Workbooks;
            var xml = @"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">";
            if (workBooks != null && workBooks.Count > 0)
            {
                for (int i = 1; i <= workBooks.Count; i++)
                {
                    var item = workBooks[i];
                    xml += $@"<menu id=""Menu{Guid.NewGuid().ToString().Replace("-", "")}""  screentip =""工作簿""   supertip=""{ConvertSpecialLetter(item.Name)}""   getImage=""GetImage""  label=""{ConvertSpecialLetter(item.Name)}"">";
                    var workSheets = item.Worksheets;
                    if (workSheets != null && workSheets.Count > 0)
                    {
                        for (int j = 1; j <= workSheets.Count; j++)
                        {
                            var sheetObj = workSheets[j];
                            var sheet = PIAHelper.GetSheetByObj(sheetObj);
                            xml += $@"<button id=""Btn{Guid.NewGuid().ToString().Replace("-", "")}""  screentip =""工作表""   supertip=""{ConvertSpecialLetter(sheet.Name)}"" getImage=""GetImage""  tag = ""{ConvertSpecialLetter(item.Name)}Separator{ConvertSpecialLetter(sheet.Name)}""  label=""{ConvertSpecialLetter(sheet.Name)}""  onAction=""Navigation""/>";
                        }
                    }
                    xml += @"</menu>";
                }
            }
            xml += @"</menu>";
            return xml;
        }

        public void Navigation(IRibbonControl control)
        {
            try
            {
                var tag = control.Tag;
                if (string.IsNullOrEmpty(tag))
                    return;
                var names = tag.Split(new[] { "Separator" }, StringSplitOptions.None);
                if (names == null || names.Count() != 2)
                    return;
                var workBookName = names[0];
                var workSheetName = names[1];
                var app = Application.GetActiveInstance();
                var workBooks = app.Workbooks;
                if (workBooks != null && workBooks.Count > 0)
                {
                    for (int i = 1; i <= workBooks.Count; i++)
                    {
                        var item = workBooks[i];
                        if (item.Name != workBookName)
                            continue;
                        else
                        {
                            item.Activate();
                            if (item.Application.ActiveWindow.WindowState == XlWindowState.xlMinimized)
                                item.Application.ActiveWindow.WindowState = XlWindowState.xlMaximized;
                            var workSheets = item.Worksheets;
                            if (workSheets != null && workSheets.Count > 0)
                            {
                                for (int j = 1; j <= workSheets.Count; j++)
                                {
                                    var sheetObj = workSheets[j];
                                    var sheet = PIAHelper.GetSheetByObj(sheetObj);
                                    if (sheet.Name != workSheetName)
                                        continue;
                                    else
                                        sheet.Select();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception) { }
        }

        string ConvertSpecialLetter(string oldString)
        {
            return oldString
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("&lt;em&gt;", "<em>")
                .Replace("&lt;/em&gt;", "</em>")
                .Replace("\"", "&quot;")
                .Replace("\'", "&apos;");
        }

        public void ChatExcelButtonClick(IRibbonControl control)
        {
            if (control.Id != MethodBase.GetCurrentMethod().Name.Replace("Click", ""))
                return;
            ProcessHelper.EnsureProcess();
            if (!LoggedIn)
            {
                var client = new NamedPipeClient<string>(NamedPipeConst.ResolveWindowPipe);
                client.PushMessage(NamedPipeConst.LoginWindow);
            }
            else
            {
                var client = new NamedPipeClient<string>(NamedPipeConst.ResolveWindowPipe);
                client.PushMessage(NamedPipeConst.ChatExcelWindow);
            }
        }
    }
}
