using System;
using System.Collections.Generic;
using static ExcelDna.Integration.Rtd.ExcelRtdServer;

namespace ChatExcel.Addin.RTD
{
    internal class RealRTDData
    {
        //topic对象
        public Topic Topic { get; set; }
        //参数
        public List<string> Params { get; set; }
        //Caller的SheetId  用于区分topic在哪个sheet下 
        public IntPtr SheetId { get; set; }
        //Value
        public object Value { get; set; }
    }
}
