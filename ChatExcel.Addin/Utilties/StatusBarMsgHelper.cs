using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChatExcel.Addin.Utilties
{
    class StatusBarMsgHelper
    {
        #region Const
        const string Default = "就绪";
        internal const string RTDStartMsg = "开始更新函数数据";
        internal const string RTDNotLoginMsg = "请确认登录信息后再试";
        internal const string RTDFinishMsg = "更新函数数据结束";
        internal const string NotLoginMsg = "未登录";
        #endregion

        static Queue<string> BarMessagesQueue = new Queue<string>();

        internal static void InitStatusBarMsg()
        {
            HandleStatusBarMsg();
        }

        private static void HandleStatusBarMsg()
        {
            Task.Factory.StartNew(async () =>
            {
                while (true)
                {
                    try
                    {
                        if (BarMessagesQueue == null || BarMessagesQueue.Count == 0)
                            await Task.Delay(200);
                        else
                        {
                            var message = BarMessagesQueue.Dequeue();
                            InvokeUtil.QueueAsMacro(() =>
                            {
                                try
                                {
                                    bool display = !string.IsNullOrEmpty(message);
                                    XlCall.Excel(XlCall.xlcMessage, display, message);
                                }
                                catch { }
                            });
                            await Task.Delay(200);
                        }
                    }
                    catch (Exception)
                    {
                        await Task.Delay(200);
                    }
                }
            });
        }

        internal static void Info(string message)
        {
            var lastOrDefault = BarMessagesQueue.LastOrDefault();
            if (string.IsNullOrEmpty(lastOrDefault))
                BarMessagesQueue.Enqueue(message);
            else if (!lastOrDefault.Equals(message))
                BarMessagesQueue.Enqueue(message);
        }

        internal static void Show(string message)
        {
            InvokeUtil.QueueAsMacro(() =>
            {
                try
                {
                    bool display = !string.IsNullOrEmpty(message);
                    XlCall.Excel(XlCall.xlcMessage, display, message);
                }
                catch { }
            });
        }
    }
}
