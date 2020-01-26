using OpenXmlSpreadsheet.Interfaces;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Reflection;

namespace OpenXmlSpreadsheet.Services
{
    public class Writer<T> : IWriter<T>
    {
        private string _SpreadSheetName = string.Empty;
        private string _WorkSheetName = string.Empty;
        private List<T> _Records = new List<T>();
        public Writer(List<T> Records)
        {
            var currentTime = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");
            string typeName = GetTypeOfT().Name.ToString();
            _SpreadSheetName = Path.Combine(string.Format("{0}_{1}.xlsx", typeName, currentTime));
            _WorkSheetName = typeName;
            _Records = Records;

        }
        public Writer(string SpreadSheetName, string SheetName, List<T> Records)
        {
            _SpreadSheetName = SpreadSheetName;
            _WorkSheetName = SheetName;
            _Records = Records;
        }
        private Type GetTypeOfT()
        {
            return typeof(T);
        }
        public void Write()
        {
            try
            {
                SpreadsheetDocument document = CreateDocument();
                using (document)
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);
                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = _WorkSheetName };
                    sheets.Append(sheet);

                    //Writing the Headers For The Column.
                    List<string> columnNames = GetColumnNames();
                    Row columnHeader = new Row();
                    foreach (var column in columnNames)
                    {
                        Cell columnHeaderCell = CreateCells(column);
                        columnHeader.Append(columnHeaderCell);
                    }
                    sheetData.AppendChild(columnHeader);

                    //Writing the Records into the Rows.
                    foreach (var record in _Records)
                    {
                        Row newRecordRow = new Row();
                        object obj = Activator.CreateInstance(record.GetType());
                        obj = record;
                        foreach (var header in columnNames)
                        {
                            Type typeOfT = GetTypeOfT();
                            PropertyInfo propertyInfo = typeOfT.GetProperty(header);
                            string value = (string)propertyInfo.GetValue(obj, null);
                            Cell recordCell = CreateCells(value);
                            newRecordRow.AppendChild(recordCell);

                        }
                        sheetData.AppendChild(newRecordRow);
                    }
                    workbookPart.Workbook.Save();
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        private SpreadsheetDocument CreateDocument()
        {

            return SpreadsheetDocument.Create(_SpreadSheetName, SpreadsheetDocumentType.Workbook);

        }
        private List<string> GetColumnNames()
        {
            List<String> columnNames = new List<string>();
            Type typeOfT = GetTypeOfT();
            PropertyInfo[] properties = typeOfT.GetProperties();
            foreach (var prop in properties)
            {
                string propName = prop.Name;
                columnNames.Add(propName);
            }
            return columnNames;
        }
        private Cell CreateCells(string CellValue)
        {
            Cell cell = new Cell();
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(CellValue);
            return cell;
        }
    }
}
