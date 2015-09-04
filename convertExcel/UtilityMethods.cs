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
        private static string GetExcelColumnName(int columnIndex)
        {
            //  Convert a zero-based column index into an Excel column reference  (A, B, C.. Y, Y, AA, AB, AC... AY, AZ, B1, B2..)
            //
            //  eg  GetExcelColumnName(0) should return "A"
            //      GetExcelColumnName(1) should return "B"
            //      GetExcelColumnName(25) should return "Z"
            //      GetExcelColumnName(26) should return "AA"
            //      GetExcelColumnName(27) should return "AB"
            //      ..etc..
            //
            if (columnIndex < 26)
                return ((char)('A' + columnIndex)).ToString();

            char firstChar = (char)('A' + (columnIndex / 26) - 1);
            char secondChar = (char)('A' + (columnIndex % 26));

            return string.Format("{0}{1}", firstChar, secondChar);
        }

        public static List<ExpandoObject> GetSpreadsheetData(string workSheet, string filePath)
        {
            List<ExpandoObject> data = new List<ExpandoObject>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                // Get the worksheet we are working with
                IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>();
                WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheets.First().Id);
                Worksheet worksheet = worksheetPart.Worksheet;
                SharedStringTablePart sstPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                SharedStringTable ssTable = null;
                if (sstPart != null)
                    ssTable = sstPart.SharedStringTable;
                // Get the CellFormats for cells without defined data types
                WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First();
                CellFormats cellFormats = (CellFormats)workbookStylesPart.Stylesheet.CellFormats;

                ExtractRowsData(data, worksheet, ssTable, cellFormats);
            }

            return data;
        }

        /// <summary>
        /// Get the data using the first row as columns and the rest of the rows as data
        /// </summary>
        /// <param name="data"></param>
        /// <param name="worksheet"></param>
        /// <param name="ssTable"></param>
        /// <param name="cellFormats"></param>
        private static void ExtractRowsData(List<ExpandoObject> data, Worksheet worksheet, SharedStringTable ssTable, CellFormats cellFormats)
        {
            var columnHeaders = worksheet.Descendants<Row>().First().Descendants<Cell>().Select(c => Convert.ToString(ProcessCellValue(c, ssTable, cellFormats))).ToArray();
            var columnHeadersCellReference = worksheet.Descendants<Row>().First().Descendants<Cell>().Select(c => c.CellReference.InnerText.Replace("1", string.Empty)).ToArray();
            // All rows are selected.
            var spreadsheetData = from row in worksheet.Descendants<Row>()
                                  select row;

            int dataRowIndex = 1;
            foreach (var dataRow in spreadsheetData)
            {
                dynamic row = new ExpandoObject();
                Cell[] rowCells = dataRow.Descendants<Cell>().ToArray();
                for (int i = 0; i < columnHeaders.Length; i++)
                {
                    // Find and add the correct cell to the row object
                    Cell cell = dataRow.Descendants<Cell>().Where(c => c.CellReference == columnHeadersCellReference[i] + dataRow.RowIndex).FirstOrDefault();
                    if (cell != null)
                        ((IDictionary<String, Object>)row).Add(new KeyValuePair<String, Object>(dataRowIndex + "," + i, ProcessCellValue(cell, ssTable, cellFormats)));
                }
                data.Add(row);
                ++dataRowIndex;
            }
        }

        /// <summary>
        /// Process the valus of a cell and return a .NET value
        /// </summary>
        static Func<Cell, SharedStringTable, CellFormats, Object> ProcessCellValue = (c, ssTable, cellFormats) =>
        {
            // If there is no data type, this must be a string that has been formatted as a number
            if (c.DataType == null && c.CellValue != null)
            {
                if (c.StyleIndex != null)
                {
                    CellFormat cf = cellFormats.Descendants<CellFormat>().ElementAt<CellFormat>(Convert.ToInt32(c.StyleIndex.Value));
                    if (cf.NumberFormatId >= 0 && cf.NumberFormatId <= 13) // This is a number
                        return Convert.ToDecimal(c.CellValue.Text);
                    else if (cf.NumberFormatId >= 14 && cf.NumberFormatId <= 22) // This is a date
                        return DateTime.FromOADate(Convert.ToDouble(c.CellValue.Text));
                    else
                        return c.CellValue.Text;
                }
            }
            else if (c.DataType == null && c.CellValue == null)
            {
                return string.Empty;
            }

            if (c.DataType != null)
            {
                switch (c.DataType.Value)
                {
                    case CellValues.SharedString:
                        return ssTable.ChildElements[Convert.ToInt32(c.CellValue.Text)].InnerText;
                    case CellValues.Boolean:
                        return c.CellValue.Text == "1" ? true : false;
                    case CellValues.Date:
                        return DateTime.FromOADate(Convert.ToDouble(c.CellValue.Text));
                    case CellValues.Number:
                        return Convert.ToDecimal(c.CellValue.Text);
                    default:
                        if (c.CellValue != null)
                            return c.CellValue.Text;
                        return string.Empty;
                }
            }
            else
            {
                if (c.CellValue != null)
                    return c.CellValue.Text;
                return string.Empty;
            }
        };

        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        // Given a WorkbookPart, inserts a new worksheet.
        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        public static void UpdateCell(string docName, string text,
            uint rowIndex, string columnName)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet =
                     SpreadsheetDocument.Open(docName, true))
            {
                WorksheetPart worksheetPart =
                      GetWorksheetPartByName(spreadSheet, "Sheet1");

                if (worksheetPart != null)
                {
                    Cell cell = GetCell(worksheetPart.Worksheet,
                                             columnName, rowIndex);

                    cell.CellValue = new CellValue(text);
                    cell.DataType =
                        new EnumValue<CellValues>(CellValues.Number);

                    // Save the worksheet.
                    worksheetPart.Worksheet.Save();
                }
            }

        }

        private static WorksheetPart
             GetWorksheetPartByName(SpreadsheetDocument document,
             string sheetName)
        {
            IEnumerable<Sheet> sheets =
               document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)
                 document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;

        }

        // Given a worksheet, a column name, and a row index, 
        // gets the cell at the specified column and row
        private static Cell GetCell(Worksheet worksheet,
                  string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().Where(c => string.Compare
                   (c.CellReference.Value, columnName +
                   rowIndex, true) == 0).First();
        }


        // Given a worksheet and a row index, return the row.
        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

        //public static List<ExpandoObject> ConvertExcelArchiveToListObjects(string filePath)
        //{
        //    DateTime begin = DateTime.UtcNow;
        //    List<ExpandoObject> listExpandoObjects = new List<ExpandoObject>();
        //    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
        //    {
        //        WorkbookPart wbPart = spreadsheetDocument.WorkbookPart;
        //        Sheets theSheets = wbPart.Workbook.Sheets;

        //        SharedStringTablePart sstPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        //        SharedStringTable ssTable = null;
        //        if (sstPart != null)
        //            ssTable = sstPart.SharedStringTable;

        //        // Get the CellFormats for cells without defined data types
        //        WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First();
        //        CellFormats cellFormats = (CellFormats)workbookStylesPart.Stylesheet.CellFormats;
        //        var sheets = wbPart.Workbook.Sheets.Cast<Sheet>().ToList();

        //        foreach (WorksheetPart worksheetpart in wbPart.WorksheetParts)
        //        {                     
        //            Worksheet worksheet = worksheetpart.Worksheet;

        //            string partRelationshipId = wbPart.GetIdOfPart(worksheetpart);
        //            var correspondingSheet = sheets.FirstOrDefault(
        //                s => s.Id.HasValue && s.Id.Value == partRelationshipId);
        //            Debug.Assert(correspondingSheet != null);
        //            string sheetName = correspondingSheet.GetAttribute("name", "").Value;
        //            // Grab the sheet name each time through your loop

        //            Debug.WriteLine(sheetName);
        //            var rowContent = worksheet.Descendants<Row>().Skip(1); 
        //            var columnHeaders = worksheet.Descendants<Row>().First().Descendants<Cell>().Select(c => Convert.ToString(ProcessCellValue(c, ssTable, cellFormats))).ToArray();

        //            dynamic expandoObjectClass = new ExpandoObject();
        //            List<ExpandoObject> listExpandoObjectRows = new List<ExpandoObject>();
        //            foreach (var dataRow in rowContent)
        //            {
        //                dynamic row = new ExpandoObject();
        //                var rowCells = dataRow.Descendants<Cell>();
        //                int cellIndex = 0;
        //                foreach (var rowCell in rowCells)
        //                {
        //                    if (rowCell.DataType != null
        //                        && rowCell.DataType.HasValue
        //                        && rowCell.DataType == CellValues.SharedString
        //                        && int.Parse(rowCell.CellValue.InnerText) < ssTable.ChildElements.Count)
        //                    {
        //                        ((IDictionary<String, Object>)row).Add(columnHeaders[cellIndex].ToString(), ssTable.ChildElements[int.Parse(rowCell.CellValue.InnerText)].InnerText ?? string.Empty);
        //                    }
        //                    else
        //                    {
        //                        if (rowCell.CellValue != null && rowCell.CellValue.InnerText != null)
        //                        {
        //                            Debug.WriteLine(rowCell.CellValue.InnerText);
        //                            ((IDictionary<String, Object>)row).Add(columnHeaders[cellIndex].ToString(), rowCell.CellValue.InnerText);
        //                        }
        //                        else 
        //                        {
        //                            Debug.WriteLine(string.Empty);
        //                            ((IDictionary<String, Object>)row).Add(columnHeaders[cellIndex].ToString(), string.Empty);
        //                        }
        //                    }
        //                    ++cellIndex;
        //                }
        //                listExpandoObjectRows.Add(row);

        //            }
        //            ((IDictionary<String, Object>)expandoObjectClass).Add(sheetName, listExpandoObjectRows);
        //            listExpandoObjects.Add(expandoObjectClass);
        //        }

        //        spreadsheetDocument.Close();
        //    }
        //    DateTime end = DateTime.UtcNow;
        //    Console.WriteLine("Measured time: " + (end-begin).TotalMinutes + " minutes.");
        //    return listExpandoObjects;
        //}

        public static List<Object> ConvertExcelArchiveToListObjects(string filePath)
        {
            DateTime begin = DateTime.UtcNow;
            List<Object> listObjects = new List<Object>();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart wbPart = spreadsheetDocument.WorkbookPart;
                Sheets theSheets = wbPart.Workbook.Sheets;

                SharedStringTablePart sstPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                SharedStringTable ssTable = null;
                if (sstPart != null)
                    ssTable = sstPart.SharedStringTable;

                // Get the CellFormats for cells without defined data types
                WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First();
                CellFormats cellFormats = (CellFormats)workbookStylesPart.Stylesheet.CellFormats;
                var sheets = wbPart.Workbook.Sheets.Cast<Sheet>().ToList();

                foreach (WorksheetPart worksheetpart in wbPart.WorksheetParts)
                {
                    Worksheet worksheet = worksheetpart.Worksheet;

                    string partRelationshipId = wbPart.GetIdOfPart(worksheetpart);
                    var correspondingSheet = sheets.FirstOrDefault(
                        s => s.Id.HasValue && s.Id.Value == partRelationshipId);
                    Debug.Assert(correspondingSheet != null);
                    // Grab the sheet name
                    string sheetName = correspondingSheet.GetAttribute("name", "").Value;

                    // create a dynamic assembly and module
                    AssemblyName assemblyName = new AssemblyName();
                    assemblyName.Name = "tmpAssembly";
                    AssemblyBuilder assemblyBuilder = Thread.GetDomain().DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
                    ModuleBuilder module = assemblyBuilder.DefineDynamicModule("tmpModule");

                    // create a new type builder
                    TypeBuilder typeBuilder = module.DefineType("MyDynamicType", TypeAttributes.Public | TypeAttributes.Class);

                    Debug.WriteLine(sheetName);
                    var rowContent = worksheet.Descendants<Row>().Skip(1);

                    var columnHeaders = worksheet.Descendants<Row>().First().Descendants<Cell>().Select(c => Convert.ToString(ProcessCellValue(c, ssTable, cellFormats))).ToArray();
                    MethodAttributes GetSetAttr =
                       MethodAttributes.Public |
                       MethodAttributes.HideBySig;
                    foreach (var columnName in columnHeaders)
                    {
                        // Generate a public property
                        var field = typeBuilder.DefineField("_" + columnName.ToString(), typeof(String), FieldAttributes.Private);
                        PropertyBuilder property =
                            typeBuilder.DefineProperty(columnName.ToString(),
                                             System.Reflection.PropertyAttributes.None,
                                             typeof(string),
                                             new Type[] { typeof(string) });

                        // Generate getter method
                        var getter = typeBuilder.DefineMethod("get_" + columnName.ToString(), GetSetAttr, typeof(String), Type.EmptyTypes);

                        var il = getter.GetILGenerator();

                        il.Emit(OpCodes.Ldarg_0);        // Push "this" on the stack
                        il.Emit(OpCodes.Ldfld, field);   // Load the field "_Name"
                        il.Emit(OpCodes.Ret);            // Return

                        property.SetGetMethod(getter);

                        // Generate setter method

                        var setter = typeBuilder.DefineMethod("set_" + columnName, GetSetAttr, null, new[] { typeof(string) });

                        il = setter.GetILGenerator();

                        il.Emit(OpCodes.Ldarg_0);        // Push "this" on the stack
                        il.Emit(OpCodes.Ldarg_1);        // Push "value" on the stack
                        il.Emit(OpCodes.Stfld, field);   // Set the field "_Name" to "value"
                        il.Emit(OpCodes.Ret);            // Return

                        property.SetSetMethod(setter);
                    }

                    dynamic expandoObjectClass = new ExpandoObject();
                    List<Object> listObjectsCustomClasses = new List<Object>();
                    foreach (var dataRow in rowContent)
                    {
                        Type generatedType = typeBuilder.CreateType();
                        object generatedObject = Activator.CreateInstance(generatedType);

                        PropertyInfo[] properties = generatedType.GetProperties();

                        int propertiesCounter = 0;

                        // Loop over the values that we will assign to the properties

                        var rowCells = dataRow.Descendants<Cell>();
                        var value = string.Empty;
                        foreach (var rowCell in rowCells)
                        {
                            if (rowCell.DataType != null
                                && rowCell.DataType.HasValue
                                && rowCell.DataType == CellValues.SharedString
                                && int.Parse(rowCell.CellValue.InnerText) < ssTable.ChildElements.Count)
                            {
                                value = ssTable.ChildElements[int.Parse(rowCell.CellValue.InnerText)].InnerText ?? string.Empty;
                            }
                            else
                            {
                                if (rowCell.CellValue != null && rowCell.CellValue.InnerText != null)
                                {
                                    value = rowCell.CellValue.InnerText;
                                }
                                else
                                {
                                    value = string.Empty;
                                }
                            }
                            properties[propertiesCounter].SetValue(generatedObject, value, null);
                            propertiesCounter++;
                        }
                        listObjectsCustomClasses.Add(generatedObject);
                    }
                    listObjects.Add(listObjectsCustomClasses);
                }
            }
            DateTime end = DateTime.UtcNow;
            Console.WriteLine("Measured time: " + (end - begin).TotalMinutes + " minutes.");
            return listObjects;
        }

        public static List<Object> ConvertExcelArchiveToListObjectsSAXApproach(string filePath)
        {
            DateTime begin = DateTime.UtcNow;
            List<Object> listObjects = new List<Object>();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart wbPart = spreadsheetDocument.WorkbookPart;
                Sheets theSheets = wbPart.Workbook.Sheets;

                SharedStringTablePart sstPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                SharedStringTable ssTable = null;
                if (sstPart != null)
                    ssTable = sstPart.SharedStringTable;

                // Get the CellFormats for cells without defined data types
                WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First();
                CellFormats cellFormats = (CellFormats)workbookStylesPart.Stylesheet.CellFormats;
                var sheets = wbPart.Workbook.Sheets.Cast<Sheet>().ToList();

                foreach (WorksheetPart worksheetpart in wbPart.WorksheetParts)
                {
                    //Worksheet worksheet = worksheetpart.Worksheet;
                    OpenXmlPartReader reader = new OpenXmlPartReader(worksheetpart);
                    bool firstRow = false;
                    String[] columnValues;
                    String rowNum;
                    Cell c;
                    var value = string.Empty;
                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(SheetData))
                        {
                            do
                            {
                                // get row child
                                reader.ReadFirstChild();
                                if (firstRow == false)
                                {
                                    firstRow = true;
                                    do
                                    {
                                        if (reader.ElementType == typeof(Row))
                                        {
                                            Debug.WriteLine("Row");
                                            reader.ReadFirstChild();
                                            do
                                            {
                                                if (reader.ElementType == typeof(Cell))
                                                {
                                                    c = (Cell)reader.LoadCurrentElement();

                                                    if (c.DataType != null
                                                        && c.DataType.HasValue
                                                        && c.DataType == CellValues.SharedString
                                                        && int.Parse(c.CellValue.InnerText) < ssTable.ChildElements.Count)
                                                    {
                                                        value = ssTable.ChildElements[int.Parse(c.CellValue.InnerText)].InnerText ?? string.Empty;
                                                    }
                                                    else
                                                    {
                                                        if (c.CellValue != null && c.CellValue.InnerText != null)
                                                        {
                                                            value = c.CellValue.InnerText;
                                                        }
                                                        else
                                                        {
                                                            value = string.Empty;
                                                        }
                                                    }
                                                    Debug.WriteLine(value);
                                                }
                                            } while (reader.ReadNextSibling()); 
                                        }

                                    } while (reader.ReadNextSibling());
                                }
                                else
                                {
                                    do
                                    {
                                        if (reader.HasAttributes)
                                        {

                                        }
                                    } while (reader.ReadNextSibling());
                                }


                            } while (reader.ReadNextSibling()); // Skip to the next row
                            break; // We just looped through all the rows so no need to continue reading the worksheet
                        }
                    }

                    //string partRelationshipId = wbPart.GetIdOfPart(worksheetpart);
                    //var correspondingSheet = sheets.FirstOrDefault(
                    //    s => s.Id.HasValue && s.Id.Value == partRelationshipId);
                    //Debug.Assert(correspondingSheet != null);
                    //// Grab the sheet name
                    //string sheetName = correspondingSheet.GetAttribute("name", "").Value;

                    //// create a dynamic assembly and module
                    //AssemblyName assemblyName = new AssemblyName();
                    //assemblyName.Name = "tmpAssembly";
                    //AssemblyBuilder assemblyBuilder = Thread.GetDomain().DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
                    //ModuleBuilder module = assemblyBuilder.DefineDynamicModule("tmpModule");

                    //// create a new type builder
                    //TypeBuilder typeBuilder = module.DefineType("MyDynamicType", TypeAttributes.Public | TypeAttributes.Class);

                    //Debug.WriteLine(sheetName);
                    //var rowContent = worksheet.Descendants<Row>().Skip(1);

                    //var columnHeaders = worksheet.Descendants<Row>().First().Descendants<Cell>().Select(c => Convert.ToString(ProcessCellValue(c, ssTable, cellFormats))).ToArray();
                    //MethodAttributes GetSetAttr =
                    //   MethodAttributes.Public |
                    //   MethodAttributes.HideBySig;
                    //foreach (var columnName in columnHeaders)
                    //{
                    //    // Generate a public property
                    //    var field = typeBuilder.DefineField("_" + columnName.ToString(), typeof(String), FieldAttributes.Private);
                    //    PropertyBuilder property =
                    //        typeBuilder.DefineProperty(columnName.ToString(),
                    //                         System.Reflection.PropertyAttributes.None,
                    //                         typeof(string),
                    //                         new Type[] { typeof(string) });

                    //    // Generate getter method
                    //    var getter = typeBuilder.DefineMethod("get_" + columnName.ToString(), GetSetAttr, typeof(String), Type.EmptyTypes);

                    //    var il = getter.GetILGenerator();

                    //    il.Emit(OpCodes.Ldarg_0);        // Push "this" on the stack
                    //    il.Emit(OpCodes.Ldfld, field);   // Load the field "_Name"
                    //    il.Emit(OpCodes.Ret);            // Return

                    //    property.SetGetMethod(getter);

                    //    // Generate setter method

                    //    var setter = typeBuilder.DefineMethod("set_" + columnName, GetSetAttr, null, new[] { typeof(string) });

                    //    il = setter.GetILGenerator();

                    //    il.Emit(OpCodes.Ldarg_0);        // Push "this" on the stack
                    //    il.Emit(OpCodes.Ldarg_1);        // Push "value" on the stack
                    //    il.Emit(OpCodes.Stfld, field);   // Set the field "_Name" to "value"
                    //    il.Emit(OpCodes.Ret);            // Return

                    //    property.SetSetMethod(setter);
                    //}

                    //dynamic expandoObjectClass = new ExpandoObject();
                    //List<Object> listObjectsCustomClasses = new List<Object>();
                    //foreach (var dataRow in rowContent)
                    //{
                    //    Type generatedType = typeBuilder.CreateType();
                    //    object generatedObject = Activator.CreateInstance(generatedType);

                    //    PropertyInfo[] properties = generatedType.GetProperties();

                    //    int propertiesCounter = 0;

                    //    // Loop over the values that we will assign to the properties

                    //    var rowCells = dataRow.Descendants<Cell>();
                    //    var value = string.Empty;
                    //    foreach (var rowCell in rowCells)
                    //    {
                    //        if (rowCell.DataType != null
                    //            && rowCell.DataType.HasValue
                    //            && rowCell.DataType == CellValues.SharedString
                    //            && int.Parse(rowCell.CellValue.InnerText) < ssTable.ChildElements.Count)
                    //        {
                    //            value = ssTable.ChildElements[int.Parse(rowCell.CellValue.InnerText)].InnerText ?? string.Empty;

                    //        }
                    //        else
                    //        {
                    //            if (rowCell.CellValue != null && rowCell.CellValue.InnerText != null)
                    //            {
                    //                value = rowCell.CellValue.InnerText;

                    //            }
                    //            else
                    //            {
                    //                value = string.Empty;
                    //            }
                    //        }
                    //        properties[propertiesCounter].SetValue(generatedObject, value, null);
                    //        propertiesCounter++;
                    //    }
                    //    listObjectsCustomClasses.Add(generatedObject);
                    //}
                    //listObjects.Add(listObjectsCustomClasses);
                }
            }
            DateTime end = DateTime.UtcNow;
            Console.WriteLine("Measured time: " + (end - begin).TotalMinutes + " minutes.");
            return listObjects;
        }
    }
}
