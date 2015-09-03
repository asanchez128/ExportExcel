using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ConvertExcel;
using System.Dynamic;

namespace ConvertExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            List<ExpandoObject> expandoList = CreateExcelFile.GetSpreadsheetData("16 SEMANAS", "carpetaPagosGrupales.xlsx");
            CreateExcelFile.WriteExcelFileFromExpandoList(expandoList, "myNewSpreadsheet.xlsx");
            CreateExcelFile.ConvertExcelArchiveToListObjects("DataDownload.xlsx");
        }
    }

    class ExportExcel
    {
        public string Email { get; set; }
    }
}
