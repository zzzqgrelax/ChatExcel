using ExcelDna.Integration;
using System;

namespace ChatExcel.Addin.Utilties
{
    public static partial class XLApp
    {
        private static bool _screenupdating = true;

        public static bool ScreenUpdating
        {
            set
            {
                XlCall.Excel(XlCall.xlcEcho, value);
                _screenupdating = value;
            }
            get
            {
                return _screenupdating;
            }
        }

        public static void ScreenUpdate(bool value)
        {
            try
            {
                if (ExcelDnaUtil.IsET)
                    return;
                bool updating = ScreenUpdating;
                if (value && !updating)
                    ScreenUpdating = true;
                if (!value && updating)
                    ScreenUpdating = false;
            }
            catch (Exception) { }
        }

        public static void ActionOnSelectedRange(this ExcelReference range, Action action)
        {
            bool updating = ScreenUpdating;

            try
            {
                if (updating) ScreenUpdating = false;

                object oldSelectionOnActiveSheet = XlCall.Excel(XlCall.xlfSelection);

                string rangeSheet = (string)XlCall.Excel(XlCall.xlSheetNm, range);

                XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { rangeSheet });
                XlCall.Excel(XlCall.xlcSelect, range);

                action.Invoke();

                XlCall.Excel(XlCall.xlcFormulaGoto, oldSelectionOnActiveSheet);
            }
            finally
            {
                if (updating) XLApp.ScreenUpdating = true;
            }
        }

        //涉及到较大数量级的操作时且操作用选中区域时使用
        //需要先设置ScreenUpdate为False
        //操作后设置ScreenUpdate为True
        public static void InvokeOnSelectedRange(this ExcelReference range, Action action)
        {
            try
            {
                string rangeSheet = (string)XlCall.Excel(XlCall.xlSheetNm, range);
                XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { rangeSheet });
                XlCall.Excel(XlCall.xlcSelect, range);
                action.Invoke();
            }
            catch
            {
                XLApp.ScreenUpdate(true);
            }
        }
    }
}
