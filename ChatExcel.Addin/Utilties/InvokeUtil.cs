using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChatExcel.Addin.Utilties
{
    public class InvokeUtil
    {
        public static void QueueAsMacro(ExcelAction action)
        {
            try
            {
                //todo...
                if (ExcelDnaUtil.IsET)
                    action.Invoke();
                else
                    ExcelAsyncUtil.QueueAsMacro(action);
            }
            catch (Exception)
            { }
        }

        public static void Run(ExcelAction action)
        {
            try
            {
                //todo..
                if (ExcelDnaUtil.IsET)
                    Task.Factory.StartNew(() => { action.Invoke(); });
                else
                    ExcelAsyncUtil.QueueAsMacro(action);
            }
            catch (Exception) { }
        }
    }
}
