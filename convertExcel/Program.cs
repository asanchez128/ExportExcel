using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ConvertExcel;
using System.Dynamic;
using System.Reflection;
using System.Diagnostics;

namespace ConvertExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // List<ExpandoObject> expandoList = CreateExcelFile.GetSpreadsheetData("16 SEMANAS", "carpetaPagosGrupales.xlsx");
            // CreateExcelFile.WriteExcelFileFromExpandoList(expandoList, "myNewSpreadsheet.xlsx");
            //CreateExcelFile.ConvertExcelArchiveToListObjects("copiaDeSample.xlsx");
            //List<List<Object>> result = CreateExcelFile.ConvertExcelArchiveToListObjectsSAXApproach("Empty.xlsx");

            //foreach (var item in result) 
            //{
            //    foreach (var subitem in item)
            //    {
            //        Type myType = subitem.GetType();
            //        IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());

            //        foreach (PropertyInfo prop in props)
            //        {
            //            object propValue = prop.GetValue(subitem, null);

            //            // Do something with propValue
            //            Debug.WriteLine(propValue);
            //        }
            //    }
            //}
            CreateExcelFile.UpdateCell("Empty.xlsx", "Hola", 1, "A");
        }
    }

    class ExportExcel
    {
        public string Email { get; set; }
    }
}
