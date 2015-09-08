using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertExcel
{
    public static class ExtensionMethods
    {
        public static Row FirstRow(this Worksheet worksheet)
        {
            Row result = new Row();
            result = worksheet.Descendants<Row>().First();
            return result;
        }

        public static Row SecondRow(this Worksheet worksheet)
        {
            Row result = new Row();
            result = worksheet.Descendants<Row>().Skip(1).First();
            return result;
        }

        public static IEnumerable<Row> SkipFirstRow(this Worksheet worksheet)
        {
            return worksheet.Descendants<Row>().Skip(1);
        }
    }
}
