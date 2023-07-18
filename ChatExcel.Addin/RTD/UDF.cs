using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using Application = NetOffice.ExcelApi.Application;

namespace ChatExcel.Addin.RTD
{
    public static class UDF
    {
        internal static Application Application;

        //Govert
        //从Excel使用基于Excel-DNA的RTD服务器有两种方法：
        //没有注册，并且包装UDF在内部调用XlCall.RTD(...)或，
        //通过注册类型（调用ExcelDna.ComInterop.ComServer.DllRegisterServer()），然后=RTD(...)直接调用该函数。
        //对于通常的第一种情况，您是对的ComVisible，RTD服务器不需要它-Excel-DNA在内部进行连接并公开类而无需注册。对于第二种情况，必须将ComVisible（显式地或通过ComVisible在类型或程序集上没有指令-因为默认值ComVisible是'true'）才能将类型注册为COM导出。
        //如果您尝试使用包装函数，则该故事会有点复杂，但是要在ProgIdExcel中进行稳定注册，以便在重新打开保存的工作表时可以使用“旧值”。
        //在这种情况下，您需要进行COM注册，并将包装器更改为call XlCall.Excel(XlCall.xlfRtd, ...)。
        //所以Excel下用XlCall.RTD   WPS下使用XlCall.Excel(XlCall.xlfRtd)较为合适

        [ExcelFunction(Description = "ChatExcel",
           Category = "ChatExcel",
           IsHidden = false,
           IsVolatile = false,
           Name = "ChatExcel")]
        public static object GetData(
    [ExcelArgument(Description = "参数1", Name = "param1", AllowReference = true)] object Param1,
    [ExcelArgument(Description = "参数2", Name = "param2", AllowReference = true)] object Param2,
    [ExcelArgument(Description = "参数3", Name = "param3", AllowReference = true)] object Param3,
    [ExcelArgument(Description = "参数4", Name = "param4", AllowReference = true)] object Param4,
    [ExcelArgument(Description = "参数5", Name = "param5", AllowReference = true)] object Param5,
    [ExcelArgument(Description = "参数6", Name = "param6", AllowReference = true)] object Param6,
    [ExcelArgument(Description = "参数7", Name = "param7", AllowReference = true)] object Param7,
    [ExcelArgument(Description = "参数8", Name = "param8", AllowReference = true)] object Param8
    )
        {
            try
            {
                var caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;

                var paramObjs = new List<object>() { Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8 };
                var topics = new List<string>();

                foreach (var item in paramObjs)
                {
                    if (item is ExcelReference)
                    {
                        var reference = item as ExcelReference;
                        var value = reference.GetValue();
                        if (value is ExcelEmpty)
                            topics.Add(string.Empty);
                        else
                            topics.Add(value.ToString());
                    }
                    else if (item is ExcelMissing)
                        continue;
                    else
                        topics.Add(item.ToString());
                }

                //加入sheetId 方便刷新
                topics.Add(caller.SheetId.ToInt64().ToString());
                var topicArray = topics.ToArray();

                if (!ExcelDnaUtil.IsET)
                    return XlCall.RTD(RtdServer.ServerProgId, null, topicArray);
                else
                {
                    try
                    {
                        object[] args = new object[topicArray.Length + 2];
                        args[0] = ComRtdServer.ServerProgId;
                        args[1] = null;
                        topicArray.CopyTo(args, 2);
                        return XlCall.Excel(XlCall.xlfRtd, args);
                    }
                    catch (Exception)
                    {
                        if (Application == null || Application.IsDisposed || Application.IsCurrentlyDisposing)
                            Application = Application.GetActiveInstance();
                        if (topicArray.Length == 0 || Application == null)
                            return ExcelError.ExcelErrorNA;
                        else if (topicArray.Length == 1)
                            return Application.WorksheetFunction.RTD(ComRtdServer.ServerProgId, null, topicArray[0]);
                        else if (topicArray.Length == 2)
                            return Application.WorksheetFunction.RTD(ComRtdServer.ServerProgId, null, topicArray[0], topicArray[1]);
                        else if (topicArray.Length == 3)
                            return Application.WorksheetFunction.RTD(ComRtdServer.ServerProgId, null, topicArray[0], topicArray[1], topicArray[2]);
                        else if (topicArray.Length == 4)
                            return Application.WorksheetFunction.RTD(ComRtdServer.ServerProgId, null, topicArray[0], topicArray[1], topicArray[2], topicArray[3]);
                        else if (topicArray.Length == 5)
                            return Application.WorksheetFunction.RTD(ComRtdServer.ServerProgId, null, topicArray[0], topicArray[1], topicArray[2], topicArray[3], topicArray[4]);
                        else if (topicArray.Length == 6)
                            return Application.WorksheetFunction.RTD(ComRtdServer.ServerProgId, null, topicArray[0], topicArray[1], topicArray[2], topicArray[3], topicArray[4], topicArray[5]);
                        else if (topicArray.Length == 7)
                            return Application.WorksheetFunction.RTD(ComRtdServer.ServerProgId, null, topicArray[0], topicArray[1], topicArray[2], topicArray[3], topicArray[4], topicArray[5], topicArray[6]);
                        else if (topicArray.Length == 8)
                            return Application.WorksheetFunction.RTD(ComRtdServer.ServerProgId, null, topicArray[0], topicArray[1], topicArray[2], topicArray[3], topicArray[4], topicArray[5], topicArray[6], topicArray[7]);
                        else if (topicArray.Length == 9)
                            return Application.WorksheetFunction.RTD(ComRtdServer.ServerProgId, null, topicArray[0], topicArray[1], topicArray[2], topicArray[3], topicArray[4], topicArray[5], topicArray[6], topicArray[7], topicArray[8]);
                        else
                            return ExcelError.ExcelErrorNA;
                    }
                }
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorGettingData;
            }
        }
    }
}
