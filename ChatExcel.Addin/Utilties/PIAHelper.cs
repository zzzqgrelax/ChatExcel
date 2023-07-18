using NetOffice.ExcelApi;
using NetOffice;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChatExcel.Addin.Utilties
{
    public class PIAHelper
    {
        public static Worksheet GetSheetByObj(object sheetobj)
        {
            if (sheetobj == null)
                return null;

            if (sheetobj is Worksheet worksheet)
                return worksheet;
            else
                return new Worksheet((ICOMObject)sheetobj);
        }

        public static Range GetRangeByObj(object rangeObj)
        {
            if (rangeObj == null)
                return null;

            if (rangeObj is Range range)
                return range;
            else
                return new Range((ICOMObject)rangeObj);
        }

        public static Areas GetAreasByObj(object areasObj)
        {
            if (areasObj == null)
                return null;

            if (areasObj is Areas areas)
                return areas;
            else
                return new Areas((ICOMObject)areasObj);
        }

        public static Worksheet GetCurrentSheet(Application app)
        {
            var sheetobj = app.ActiveSheet;
            return GetSheetByObj(sheetobj);
        }
    }
}
