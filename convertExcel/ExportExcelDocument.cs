using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertExcel
{
    public partial class CreateExcelFile
    {
        ///// <summary>
        ///// Abre un archivo xlsx y extrae la información de una hoja de excel para convertirla 
        ///// en una lista de objetos expandos usando la función 
        ///// ExtractRowsData(data, worksheet, ssTable, cellFormats);
        ///// </summary>
        ///// <param name="workSheet"> Índice cero de la columna de Excel. </param>
        ///// <returns> Una lista con objetos expandos. </returns>
        ///// <exception cref="System.IO.IOException">Excepción lanzada cuando el archivo de Excel 
        ///// está siendo usado por otro proceso.</exception>
        //public static List<ExpandoObject> GetSpreadsheetData(string workSheet, string filePath)
        //{
        //    List<ExpandoObject> data = new List<ExpandoObject>();

        //    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
        //    {
        //        IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>();
        //        WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheets.First().Id);
        //        Worksheet worksheet = worksheetPart.Worksheet;
        //        SharedStringTablePart sstPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        //        SharedStringTable ssTable = null;
        //        if (sstPart != null)
        //            ssTable = sstPart.SharedStringTable;

        //        WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First();
        //        CellFormats cellFormats = (CellFormats)workbookStylesPart.Stylesheet.CellFormats;

        //        ExtractRowsData(data, worksheet, ssTable, cellFormats);
        //    }

        //    return data;
        //}

        /// <summary>
        /// Obtener informacion de una hoja de Excel asumiendo que en la primera fila se nombran las columnas 
        /// y que el resto de las filas contiene información asociada al nombre de las columnas
        /// </summary>
        /// <param name="data">La lista de objetos expando que guardará toda la información que se extraiga
        /// de la tabla.</param>
        /// <param name="worksheet">El nombre de la hoja de Excel.</param>
        /// <param name="ssTable">La tabla de cadenas compartidas del archivo xlsx en el cual se encuentra 
        /// la hoja de Excel de la que se extraerán los datos.</param>
        /// <param name="cellFormats">El formato de celdas.</param>
        private static void ExtractRowsData(List<ExpandoObject> data, Worksheet worksheet, SharedStringTable ssTable, CellFormats cellFormats)
        {
            var columnHeaders = worksheet.FirstRow().Descendants<Cell>().Select(c => Convert.ToString(ProcessCellValue(c, ssTable, cellFormats))).ToArray();
            var columnHeadersCellReference = worksheet.FirstRow().Descendants<Cell>().Select(c => c.CellReference.InnerText.Replace("1", string.Empty)).ToArray();

            var spreadsheetData = worksheet.SkipFirstRow();

            int dataRowIndex = 2;
            foreach (var dataRow in spreadsheetData)
            {
                dynamic row = new ExpandoObject();
                Cell[] rowCells = dataRow.Descendants<Cell>().ToArray();
                for (int i = 0; i < columnHeaders.Length; i++)
                {
                    // Selecciona y agrega la celda correcta al archivo de la fila.
                    Cell cell = dataRow.Descendants<Cell>().Where(c => c.CellReference == columnHeadersCellReference[i] + dataRow.RowIndex).FirstOrDefault();
                    if (cell != null)
                        ((IDictionary<String, Object>)row).Add(new KeyValuePair<String, Object>(dataRowIndex + "," + i, ProcessCellValue(cell, ssTable, cellFormats)));
                }
                data.Add(row);
                ++dataRowIndex;
            }
        }

        private static void ExtractRowsData(List<Object> data, Worksheet worksheet, SharedStringTable ssTable, CellFormats cellFormats)
        {
            var columnHeaders = worksheet.FirstRow().Descendants<Cell>().Select(c => Convert.ToString(ProcessCellValue(c, ssTable, cellFormats))).ToArray();
            var columnHeadersCellReference = worksheet.FirstRow().Descendants<Cell>().Select(c => c.CellReference.InnerText.Replace("1", string.Empty)).ToArray();

            var spreadsheetData = worksheet.SkipFirstRow();

            int dataRowIndex = 2;
            foreach (var dataRow in spreadsheetData)
            {
                dynamic row = new ExpandoObject();
                Cell[] rowCells = dataRow.Descendants<Cell>().ToArray();
                for (int i = 0; i < columnHeaders.Length; i++)
                {
                    // Selecciona y agrega la celda correcta al archivo de la fila.
                    Cell cell = dataRow.Descendants<Cell>().Where(c => c.CellReference == columnHeadersCellReference[i] + dataRow.RowIndex).FirstOrDefault();
                    if (cell != null)
                        ((IDictionary<String, Object>)row).Add(new KeyValuePair<String, Object>(dataRowIndex + "," + i, ProcessCellValue(cell, ssTable, cellFormats)));
                }
                data.Add(row);
                ++dataRowIndex;
            }
        }
        /// <summary>
        /// Procesa los valores de la celda y regresa un valor .NET
        /// </summary>
        static Func<Cell, SharedStringTable, CellFormats, Object> ProcessCellValue = (c, ssTable, cellFormats) =>
        {
            // Si no hay tipo de datos, debe de ser una cadena que ha sido formateada como número.
            if (c.DataType == null && c.CellValue != null)
            {
                if (c.StyleIndex != null)
                {
                    CellFormat cf = cellFormats.Descendants<CellFormat>().ElementAt<CellFormat>(Convert.ToInt32(c.StyleIndex.Value));
                    if (cf.NumberFormatId >= 0 && cf.NumberFormatId <= 13) // Es un número
                        return Convert.ToDecimal(c.CellValue.Text);
                    else if (cf.NumberFormatId >= 14 && cf.NumberFormatId <= 22) // Es una fecha.
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

        /// <summary>
        /// Convierte el contenido de un archivo de Excel en una lista de objectos usando un parseador de XML.
        /// Asume que cada hoja en el archivo de Excel representa a un objecto y que la primera fila de cada 
        /// hoja contiene las columnas que describen las propiedades del objeto. También asume que los nombres 
        /// de las columnas no se repiten. En caso de repetirse el nombre de alguna columna solo se guardará en 
        /// el objeto una sola vez.
        /// También se da por sentado que todos los datos en una columna a excepción del encabezado son del 
        /// mismo tipo.
        /// </summary>
        /// <exception cref="System.IO.IOException">Excepción lanzada cuando el archivo de Excel 
        /// está siendo usado por otro proceso.</exception>

        public static List<List<Object>> ConvertExcelArchiveToListObjectsSAXApproach(string filePath)
        {
            List<List<Object>> listObjects = new List<List<Object>>();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart wbPart = spreadsheetDocument.WorkbookPart;
                Sheets theSheets = wbPart.Workbook.Sheets;

                SharedStringTablePart sstPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                SharedStringTable ssTable = null;
                if (sstPart != null)
                    ssTable = sstPart.SharedStringTable;

                // Obtiene el formato de celdas para las celdas sin tipos de datos definidos.
                WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First();
                CellFormats cellFormats = (CellFormats)workbookStylesPart.Stylesheet.CellFormats;
                var sheets = wbPart.Workbook.Sheets.Cast<Sheet>().ToList();

                foreach (WorksheetPart worksheetpart in wbPart.WorksheetParts)
                {
                    OpenXmlPartReader reader = new OpenXmlPartReader(worksheetpart);
                    List<String> columnValues = new List<String>();
                    Cell c;
                    var value = string.Empty;

                    string partRelationshipId = wbPart.GetIdOfPart(worksheetpart);
                    var correspondingSheet = sheets.FirstOrDefault(
                        s => s.Id.HasValue && s.Id.Value == partRelationshipId);

                    string sheetName = string.Empty;

                    // Obtiene el nombre de la hoja de Excel
                    if (correspondingSheet != null)
                    {
                        sheetName = correspondingSheet.GetAttribute("name", "").Value;
                    }

                    if (sheetName == string.Empty)
                    {
                        sheetName = "MyDynamicType";
                    }
                    // Crea un ensamblado dinámico y un módulo.
                    AssemblyName assemblyName = new AssemblyName();
                    assemblyName.Name = "tmpAssembly";
                    AssemblyBuilder assemblyBuilder = Thread.GetDomain().DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
                    ModuleBuilder module = assemblyBuilder.DefineDynamicModule("tmpModule");

                    // Crea un nuevo constructor de tipos.
                    TypeBuilder typeBuilder = module.DefineType(sheetName, TypeAttributes.Public | TypeAttributes.Class);

                    MethodAttributes GetSetAttr =
                       MethodAttributes.Public |
                       MethodAttributes.HideBySig;
                    int numberOfColumns = 0;

                    bool firstRowInformation = false;

                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(SheetData))
                        {
                            // Obtiene la primera fila.
                            // Asume que no hay celdas en blanco entre las columnas con información.
                            // Si hay columnas vacias no serán agregadas. 
                            reader.ReadFirstChild();
                            if (reader.ElementType == typeof(Row))
                            {
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
                                        if (value != string.Empty)
                                        {
                                            columnValues.Add(value);
                                            ++numberOfColumns;
                                        }
                                    }
                                } while (reader.ReadNextSibling());
                            }
                            reader.Read();

                            foreach (var columnName in columnValues)
                            {
                                // Genera una propiedad pública
                                var field = typeBuilder.DefineField("_" + columnName.ToString(), typeof(String), FieldAttributes.Private);
                                PropertyBuilder property =
                                    typeBuilder.DefineProperty(columnName.ToString(),
                                                        System.Reflection.PropertyAttributes.None,
                                                        typeof(string),
                                                        new Type[] { typeof(string) });

                                // Genera un método para acceder a la propiedad.
                                var getter = typeBuilder.DefineMethod("get_" + columnName.ToString(), GetSetAttr, typeof(String), Type.EmptyTypes);

                                var il = getter.GetILGenerator();

                                il.Emit(OpCodes.Ldarg_0);
                                il.Emit(OpCodes.Ldfld, field);
                                il.Emit(OpCodes.Ret);

                                property.SetGetMethod(getter);

                                // Genera un método para cambiar el valor de la propiedad. 
                                var setter = typeBuilder.DefineMethod("set_" + columnName, GetSetAttr, null, new[] { typeof(string) });

                                il = setter.GetILGenerator();

                                il.Emit(OpCodes.Ldarg_0);
                                il.Emit(OpCodes.Ldarg_1);
                                il.Emit(OpCodes.Stfld, field);
                                il.Emit(OpCodes.Ret);

                                property.SetSetMethod(setter);
                            }

                            dynamic expandoObjectClass = new ExpandoObject();
                            List<Object> listObjectsCustomClasses = new List<Object>();

                            // Lee el resto de las filas en la hoja de Excel
                            do
                            {
                                if (reader.ElementType == typeof(Row))
                                {
                                    reader.ReadFirstChild();
                                    Type generatedType = typeBuilder.CreateType();
                                    object generatedObject = Activator.CreateInstance(generatedType);

                                    PropertyInfo[] properties = generatedType.GetProperties();

                                    int propertiesCounter = 0;
                                    // Lee todas las celdas en la fila.

                                    if (firstRowInformation == false)
                                    {
                                        firstRowInformation = true;
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

                                                if (propertiesCounter < properties.Count())
                                                {
                                                    properties[propertiesCounter].SetValue(generatedObject, value, null);
                                                }
                                                propertiesCounter++;
                                            }
                                        } while (reader.ReadNextSibling());
                                    }

                                    else
                                    {
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

                                                if (propertiesCounter < properties.Count())
                                                {
                                                    if (value == string.Empty && properties[propertiesCounter].GetType().Name != "String") 
                                                    {
                                                        properties[propertiesCounter].SetValue(generatedObject, null, null);
                                                    }
                                                    else
                                                    {
                                                        properties[propertiesCounter].SetValue(generatedObject, value, null);
                                                    }
                                                    
                                                }
                                                propertiesCounter++;
                                            }
                                        } while (reader.ReadNextSibling());
                                    }
                                    listObjectsCustomClasses.Add(generatedObject);
                                }
                            } while (reader.Read() && reader.ElementType == typeof(Row));
                            listObjects.Add(listObjectsCustomClasses);
                        }
                    }
                }
            }
            return listObjects;
        }

        public static List<List<Object>> ConvertExcelArchiveToListObjects(string filePath)
        {
            List<List<Object>> listObjects = new List<List<Object>>();
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
                    string sheetName = string.Empty;
                    sheetName = correspondingSheet.GetAttribute("name", "").Value;

                    if (sheetName == string.Empty)
                    {
                        sheetName = "MyDynamicType";
                    }

                    var columnHeaders = worksheet.FirstRow().
                        Descendants<Cell>().Select(c => Convert.ToString(ProcessCellValue(c, ssTable, cellFormats)));

                    var secondRowTypes = worksheet.SecondRow().
                        Descendants<Cell>().Select(c => Type.GetType("System." + Convert.GetTypeCode(ProcessCellValue(c, ssTable, cellFormats))));

                    var listTypes = columnHeaders.Zip(secondRowTypes, (c, r) => new Field
                        (
                            c,
                            r
                        )
                    );

                    var rowContent = worksheet.SkipFirstRow();

                    Type generatedType = TypeGenerator.CompileResultType(sheetName, listTypes);
                    
                    dynamic expandoObjectClass = new ExpandoObject();
                    List<Object> listObjectsCustomClasses = new List<Object>();
                    foreach (var dataRow in rowContent)
                    {
                        var generatedObject = Activator.CreateInstance(generatedType);
                        PropertyInfo[] properties = generatedType.GetProperties();
                        int propertiesCounter = 0;

                        // Loop over the values that we will assign to the properties
                        var rowCells = dataRow.Descendants<Cell>();
                        
                        foreach (var rowCell in rowCells)
                        {
                            var value = ProcessCellValue(rowCell, ssTable, cellFormats);
                            if (propertiesCounter < properties.Count()) 
                            {
                                if (value.GetType().Name == "String" 
                                    && value == string.Empty 
                                    && properties[propertiesCounter].GetType().Name != "String")
                                {
                                    properties[propertiesCounter].SetValue(generatedObject, null, null);
                                }
                                else
                                {
                                    properties[propertiesCounter].SetValue(generatedObject, value, null);
                                }
                            }
                            
                            propertiesCounter++;
                        }
                        listObjectsCustomClasses.Add(generatedObject);
                    }
                    listObjects.Add(listObjectsCustomClasses);
                }
            }
            return listObjects;
        }
    }

    public sealed class Field
    {
        private readonly String fieldName;
        private readonly Type fieldType;
        public String FieldName { get { return fieldName; } }
        public Type FieldType { get { return fieldType; } }

        public Field(String fName, Type fType)
        {
            this.fieldName = fName;
            this.fieldType = fType;
        }
    }
}
