using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Data;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Dynamic;
using System.Reflection.Emit;
using System.Threading;

namespace ConvertExcel
{
    public partial class CreateExcelFile
    {
        public static bool CreateExcelDocument<T>(List<T> list, string xlsxFilePath)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(ListToDataTable(list));

            return CreateExcelDocument(ds, xlsxFilePath);
        }

        public static bool CreateExcelDocument(DataTable dt, string xlsxFilePath)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            bool result = CreateExcelDocument(ds, xlsxFilePath);
            ds.Tables.Remove(dt);
            return result;
        }

        /// <summary>
        /// Create an Excel file, and write it out to a MemoryStream (rather than directly to a file)
        /// </summary>
        /// <param name="dt">DataTable containing the data to be written to the Excel.</param>
        /// <param name="filename">The filename (without a path) to call the new Excel file.</param>
        /// <param name="Response">HttpResponse of the current page.</param>
        /// <returns>True if it was created succesfully, otherwise false.</returns>
        public static bool CreateExcelDocument(DataTable dt)
        {
            try
            {
                DataSet ds = new DataSet();
                ds.Tables.Add(dt);
                CreateExcelDocumentAsStream(ds);
                ds.Tables.Remove(dt);
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Failed, exception thrown: " + ex.Message);
                return false;
            }
        }

        public static bool CreateExcelDocument<T>(List<T> list)
        {
            try
            {
                DataSet ds = new DataSet();
                ds.Tables.Add(ListToDataTable(list));
                CreateExcelDocumentAsStream(ds);
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Failed, exception thrown: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Create an Excel file, and write it out to a MemoryStream (rather than directly to a file)
        /// </summary>
        /// <param name="ds">DataSet containing the data to be written to the Excel.</param>
        /// <param name="filename">The filename (without a path) to call the new Excel file.</param>
        /// <param name="Response">HttpResponse of the current page.</param>
        /// <returns>Either a MemoryStream, or NULL if something goes wrong.</returns>
        public static byte[] CreateExcelDocumentAsStream(DataSet ds)
        {
            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true))
            {
                WriteExcelFile(ds, document);
            }
            stream.Flush();
            stream.Position = 0;

            byte[] data1 = new byte[stream.Length];
            stream.Read(data1, 0, data1.Length);
            stream.Close();
            return data1;
        }


        /// <summary>
        /// Create an Excel file, and write it out to a MemoryStream (rather than directly to a file)
        /// </summary>
        /// <param name="ds">DataSet containing the data to be written to the Excel.</param>
        /// <param name="filename">The filename (without a path) to call the new Excel file.</param>
        /// <param name="Response">HttpResponse of the current page.</param>
        /// <returns>Either a MemoryStream, or NULL if something goes wrong.</returns>
        public static byte[] CreateExcelDocumentAsStream(List<List<Object>> listObjects)
        {
            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true))
            {
                WriteListListObjectsToExcelFile(listObjects, document);
            }
            stream.Flush();
            stream.Position = 0;

            byte[] data1 = new byte[stream.Length];
            stream.Read(data1, 0, data1.Length);
            stream.Close();
            return data1;
        }

        #region public static bool CreateExcelDocument(DataSet ds, string excelFilename)
        /// <summary>
        /// Create an Excel file, and write it to a file.
        /// </summary>
        /// <param name="ds">DataSet containing the data to be written to the Excel.</param>
        /// <param name="excelFilename">Name of file to be written.</param>
        /// <returns>True if successful, false if something went wrong.</returns>
        public static bool CreateExcelDocument(DataSet ds, string excelFilename)
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(excelFilename, SpreadsheetDocumentType.Workbook))
                {
                    WriteExcelFile(ds, document);
                }
                Trace.WriteLine("Successfully created: " + excelFilename);
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Failed, exception thrown: " + ex.Message);
                return false;
            }
        }
        #endregion

        #region private static void WriteExcelFile(DataSet ds, SpreadsheetDocument spreadsheet)
        private static void WriteExcelFile(DataSet ds, SpreadsheetDocument spreadsheet)
        {
            //  Create the Excel file contents.  This function is used when creating an Excel file either writing 
            //  to a file, or writing to a MemoryStream.
            WorkbookPart workbookpart = spreadsheet.AddWorkbookPart();
            workbookpart.Workbook = new Workbook(); 
            workbookpart.Workbook.Save();
            WorkbookPart wbPart = spreadsheet.WorkbookPart;

            spreadsheet.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

            WorkbookStylesPart workbookStylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
            Stylesheet stylesheet = new Stylesheet();
            workbookStylesPart.Stylesheet = stylesheet;

            SharedStringTablePart shareStringPart;
            if (workbookpart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = workbookpart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = workbookpart.AddNewPart<SharedStringTablePart>();
            }

            uint worksheetNumber = 1;

            foreach (DataTable dt in ds.Tables)
            {
                //  For each worksheet you want to create
                string workSheetID = "rId" + worksheetNumber.ToString();
                string worksheetName = dt.TableName;

                WorksheetPart newWorksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();

                // create sheet data
                newWorksheetPart.Worksheet.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

                // save worksheet
                WriteDataTableToExcelWorksheet(dt, newWorksheetPart);
                newWorksheetPart.Worksheet.Save();

                // create the worksheet to workbook relation
                if (worksheetNumber == 1)
                    spreadsheet.WorkbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());

                spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                {
                    Id = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart),
                    SheetId = (uint)worksheetNumber,
                    Name = dt.TableName
                });

                worksheetNumber++;
            }

            spreadsheet.WorkbookPart.Workbook.Save();
        }
        #endregion

        private static void WriteListListObjectsToExcelFile(List<List<Object>> listObjects, SpreadsheetDocument spreadsheet)
        {
            WorkbookPart workbookpart = spreadsheet.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            WorkbookPart wbPart = spreadsheet.WorkbookPart;

            spreadsheet.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

            WorkbookStylesPart workbookStylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
            Stylesheet stylesheet = new Stylesheet();
            workbookStylesPart.Stylesheet = stylesheet;

            //SharedStringTablePart shareStringPart;
            //if (workbookpart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            //{
            //    shareStringPart = workbookpart.GetPartsOfType<SharedStringTablePart>().First();
            //}
            //else
            //{
            //    shareStringPart = workbookpart.AddNewPart<SharedStringTablePart>();
            //}
            
            //  Loop through each of the DataTables in our DataSet, and create a new Excel Worksheet for each.
            uint worksheetNumber = 1;
            


            foreach (var item in listObjects)
            {
                string workSheetID = "rId" + worksheetNumber.ToString();
                string worksheetName = string.Empty;
                if (listObjects.Count > 0)
                {
                    worksheetName = listObjects[0].GetType().Name.ToString();
                }
                if (worksheetName == string.Empty)
                {
                    worksheetName = worksheetNumber.ToString();
                }

                WorksheetPart newWorksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();

                // create sheet data
                newWorksheetPart.Worksheet.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

                // save worksheet
                WriteListObjectsToExcelWorksheet(item, newWorksheetPart);
                newWorksheetPart.Worksheet.Save();

                // create the worksheet to workbook relation
                if (worksheetNumber == 1)
                    spreadsheet.WorkbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());

                spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                {
                    Id = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart),
                    SheetId = (uint)worksheetNumber,
                    Name = worksheetNumber.ToString()
                });

                worksheetNumber++;
            }

            spreadsheet.WorkbookPart.Workbook.Save();
        }

        #region WriteExcelFileFromExpandoList
        public static void WriteExcelFileFromExpandoList(List<ExpandoObject> expandoListObject, string filePath)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            uint row = 0, col = 0;

            foreach (var item in expandoListObject)
            {
                IDictionary<string, object> expDict = item;

                foreach (var keyValue in expDict)
                {
                    var keyValueArray = keyValue.Key.Split(',');
                    row = (uint)Int32.Parse(keyValueArray[0]);
                    col = (uint)Int32.Parse(keyValueArray[1]);
                    Cell cell = InsertCellInWorksheet(GetExcelColumnName((int) col), row, worksheetPart);
                    cell.CellValue = new CellValue(keyValue.Value.ToString());
                    cell.DataType = CellValues.String;
                }
            }
            spreadsheetDocument.Close();
        }
        #endregion

        

        private static void WriteListObjectsToExcelWorksheet(List<Object> listObject, WorksheetPart worksheetPart)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            uint rowIndex = 1;
            int propertyCounter = 0;
            List<String> excelColumnNames = new List<String>();
            
            if (listObject.Count > 0) 
            {
                var headerRow = new Row { RowIndex = rowIndex };  // add a row at the top of spreadsheet
                sheetData.Append(headerRow);

                Type myType = listObject[0].GetType();
                IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());
                foreach (var property in props) 
                {
                    excelColumnNames.Add(GetExcelColumnName(propertyCounter));
                    AppendTextCell(GetExcelColumnName(propertyCounter) + "1", property.Name.ToString(), headerRow);
                    ++propertyCounter;
                }
            }
            
            foreach (var subitem in listObject)
            {
                Type myType = subitem.GetType();
                IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());

                ++rowIndex;
                var newExcelRow = new Row { RowIndex = rowIndex };  // add a row at the top of spreadsheet
                sheetData.Append(newExcelRow);
                int colInx = 0;
                foreach (PropertyInfo prop in props)
                {
                    object propValue = prop.GetValue(subitem, null);
                    AppendTextCell(excelColumnNames[colInx] + rowIndex.ToString(), propValue.ToString(), newExcelRow);
                    
                    ++colInx;
                    Debug.WriteLine(propValue);
                }
            }
        }
    }
}
