using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EZExcel
{
    // This API is an wrapper class of OpenXml.
    // This API allows users to easily and efficiently use OpenXml with less background of it.
    public class Excel
    {
        #region APIs for reading Excel
        public static SpreadsheetDocument Open(string fileName)
        {
            SpreadsheetDocument document = null;

            if (File.Exists(fileName) == true)
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    document = SpreadsheetDocument.Open(fileStream, false);
                }
            }
            
            return document;
        }

        public static Worksheet GetWorksheet(SpreadsheetDocument document, int worksheetIndex = 0)
        {
            Sheet sheet = document.WorkbookPart.Workbook.Sheets.ChildElements[worksheetIndex] as Sheet;
            return (document.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
        }

        public static Worksheet GetWorksheet(string fileName, int worksheetIndex = 0)
        {
            SpreadsheetDocument document = Open(fileName);
            if (document == null)
            {
                return null;
            }

            return GetWorksheet(document, worksheetIndex);
        }

        public static int GetWorksheetCount(SpreadsheetDocument document)
        {
            return document.WorkbookPart.WorksheetParts.Count();
        }

        public static List<Row> GetRows(Worksheet worksheet)
        {
            return worksheet.GetFirstChild<SheetData>().Descendants<Row>().ToList();
        }

        public static int GetRowCount(Worksheet worksheet)
        {
            return worksheet.GetFirstChild<SheetData>().Descendants<Row>().ToList().Count;
        }

        public static List<Cell> GetCells(Row row)
        {
            return row.Descendants<Cell>().ToList();
        }

        public static int GetCellCount(Row row)
        {
            return row.Descendants<Cell>().ToList().Count;
        }

        public static string GetCellData(SpreadsheetDocument document, Cell cell)
        {
            if (cell.CellValue == null)
            {
                return string.Empty;
            }

            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return document.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            }
            return value;
        }
        #endregion

        #region APIs for writing Excel
        public static SpreadsheetDocument Create(string fileName, bool createNewWorksheet = true)
        {
            SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
            document.AddWorkbookPart().Workbook = new Workbook();
            
            if (createNewWorksheet == true)
            {
                AddWorksheet(document);
            }

            document.WorkbookPart.Workbook.Save();

            return document;
        }

        public static Worksheet AddWorksheet(SpreadsheetDocument document, string worksheetName = "Sheet")
        {
            return InsertWorksheet(document.WorkbookPart, worksheetName);
        }

        // Indexes (rowIndex and columnIndex) start at 1.
        // If they are less than 1, then it returns null.
        public static Cell WriteCell(Worksheet worksheet, int rowIndex, int columnIndex, string data)
        {
            if (rowIndex < 1 || columnIndex < 1)
            {
                return null;
            }

            return InsertCellInWorksheet(worksheet, rowIndex, ConvertToColumnString(columnIndex), data);
        }

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
        private static Worksheet InsertWorksheet(WorkbookPart workbookPart, string worksheetName = "Sheet")
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets == null)
            {
                sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            }

            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = worksheetName + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart.Worksheet;
        }

        // Given a column name, a row index, and a Worksheet, inserts a cell into the worksheet. 
        // If the cell already exists, overwrites and returns it. 
        private static Cell InsertCellInWorksheet(Worksheet worksheet, int rowIndex, string columnName, string data)
        {
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
                row = new Row() { RowIndex = Convert.ToUInt32(rowIndex) };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                Cell cell = row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
                cell.CellValue = new CellValue(data); // overwrite
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                return cell;
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
                newCell.CellValue = new CellValue(data);
                newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        private static string ConvertToColumnString(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static void Save(Worksheet worksheet)
        {
            if (worksheet != null)
            {
                worksheet.Save();
            }
        }
        #endregion

        #region Common
        public static void Close(SpreadsheetDocument document)
        {
            if (document != null)
            {
                document.Close();
            }
        }
        #endregion
    }
}
