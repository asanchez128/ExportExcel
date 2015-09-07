using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ConvertExcel;
using System.Dynamic;
using System.Reflection;
using System.Diagnostics;
using System.IO;

namespace ConvertExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // List<ExpandoObject> expandoList = CreateExcelFile.GetSpreadsheetData("16 SEMANAS", "carpetaPagosGrupales.xlsx");
            // CreateExcelFile.WriteExcelFileFromExpandoList(expandoList, "myNewSpreadsheet.xlsx");
            //CreateExcelFile.ConvertExcelArchiveToListObjects("copiaDeSample.xlsx");
            List<List<Object>> result = CreateExcelFile.ConvertExcelArchiveToListObjectsSAXApproach("Empty.xlsx");

            
            CreateExcelFile.UpdateCell("Empty.xlsx", "Hola", 1, "A");
            File.WriteAllBytes("Test1.xlsx", CreateExcelFile.CreateExcelDocumentAsStream(result));
            //CreateExcelFile.CreateExcelDocumentAsStream(result);
        }
    }

    class ExportExcel
    {
        public string Email { get; set; }
    }
}
