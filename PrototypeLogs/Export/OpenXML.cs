using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Pathway.WPF.ImportExport.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrototypeLogs.Export
{
    public class OpenXML: ExcelBuilderOpenXML
    {
        private string _fileName;
        private SpreadsheetDocument _document;
        protected WorkbookPart book;
        protected Worksheet Worksheet;


        public WorksheetPart Sheet;
        public SheetData SheetData;
        private uint rowIdx = 0;
        private int columnStepIndex = 1;

        public SpreadsheetDocument Document => spreadsheetDocument;
        /// <summary>
        /// Creates new OpenXML
        /// </summary>
        /// <param name="fileName">Name of file to export in</param>
        public OpenXML(string fileName, string sheetName, uint sheetIndex, bool isReadonly = false):base(fileName)
        {
            if (String.IsNullOrEmpty(fileName))
                throw new ArgumentException("fileName");

            //if (File.Exists(fileName))
            //    _document = SpreadsheetDocument.Open(fileName, isReadonly);
            //else
            //    _document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

            KeyValuePair<WorksheetPart, SheetData> dataSheets = CreateSheet(sheetName, sheetIndex);
            Sheet = dataSheets.Key;
            SheetData = dataSheets.Value;
            _fileName = fileName;
        }

        public OpenXML(string fileName, string sheetName, uint sheetIndex, bool withReopen, bool isReadonly = false) : base(fileName, true)
        {
            if (String.IsNullOrEmpty(fileName))
                throw new ArgumentException("fileName");

            //if (File.Exists(fileName))
            //    _document = SpreadsheetDocument.Open(fileName, isReadonly);
            //else
            //    _document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

            KeyValuePair<WorksheetPart, SheetData> dataSheets = CreateSheet(sheetName, sheetIndex);
            Sheet = dataSheets.Key;
            SheetData = dataSheets.Value;
            _fileName = fileName;
        }

        public int CountSheets => _document.WorkbookPart.Workbook.Sheets.Count();

        public Worksheet GetworksheetBySheetName(string sheetName)
        {
            var workbookPart = spreadsheetDocument.WorkbookPart;
            string relationshipId = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name.Equals(sheetName))?.Id;

            var worksheet = ((WorksheetPart)workbookPart.GetPartById(relationshipId)).Worksheet;

            return worksheet;
        }

        /// <summary>
		/// Adds header
		/// </summary>
		/// <param name="title">Title</param>
		public void AddHeader(string title)
        {
            if (String.IsNullOrEmpty(title))
                throw new ArgumentException("title");

            if (Sheet != null && SheetData != null)
            {
                var row = SheetData.AppendChild(new Row());
                InsertCell(row, title, CellValues.String, HEADER1INDEXSTYLE);
                InsertCell(row, string.Empty, CellValues.String, HEADER1INDEXSTYLE);
            }
        }

        /// <summary>
		/// Adds from dictionary<int, string>
		/// </summary>
		/// <param name="title">Title</param>
		public void AddHeader(Dictionary<int, string> header)
        {
            if (header == null)
                throw new ArgumentException("Header dictionary for export logs to excel file is empty");

            if (Sheet != null && SheetData != null)
            {
                var row = SheetData.AppendChild(new Row());
                foreach (KeyValuePair<int, string> item in header) {
                    InsertCell(row, item.Value, CellValues.String, HEADER1INDEXSTYLE);
                }

                //InsertCell(row, title, CellValues.String, HEADER1INDEXSTYLE);
                //InsertCell(row, string.Empty, CellValues.String, HEADER1INDEXSTYLE);
            }
        }

        public void SetCurrentSheetByName(string name) {
            var sheet = GetworksheetBySheetName(name);
            Worksheet = sheet;
        }

        //public Sheet GetSheet(string name) {
        //    return spreadsheetDocument.WorkbookPart.Workbook.Sheets.Where(w => string.CompareOrdinal(w.LocalName, name) == 0).SingleOrDefault();
        //}

        /// </summary>
        /// <param name="row">Row</param>
        /// <param name="value">string</param>
        /// <param name="dateType">CellValues : Default CellValues.String</param>
        /// <param name="styleIndex">styleIndex : Default = 1; Bold = 2; Header1 = 3</param>
        public void InsertCell(Row row, string value, CellValues dateType = CellValues.String, uint styleIndex = ExcelConstant.DEFAULTINDEXSTYLE)
        {
            row.AppendChild(new Cell() { CellValue = new CellValue(value), DataType = dateType, StyleIndex = styleIndex });
        }

        public void InsertColumn(Row row, string value, CellValues dateType = CellValues.String, uint styleIndex = ExcelConstant.DEFAULTINDEXSTYLE)
        {
            var c = new Column() { };
            row.AppendChild(new Cell() { CellValue = new CellValue(value), DataType = dateType, StyleIndex = styleIndex });
        }


        public void Close()
        {
            spreadsheetDocument.WorkbookPart.Workbook.Save();
            spreadsheetDocument.Close();
        }

        public void SetColumnWidth(int from, int to, int width)
        {
            // Save the stylesheet formats
            //stylesPart.Stylesheet.Save();
            
            // Create custom widths for columns
            Columns lstColumns = Sheet.Worksheet.GetFirstChild<Columns>();
            Boolean needToInsertColumns = false;
            if (lstColumns == null)
            {
                lstColumns = new Columns();
                needToInsertColumns = true;
            }
            //Sheet.Worksheet.
            // Min = 1, Max = 1 ==> Apply this to column 1 (A)
            // Min = 2, Max = 2 ==> Apply this to column 2 (B)
            // Width = 25 ==> Set the width to 25
            // CustomWidth = true ==> Tell Excel to use the custom width
            //for (var i = from; i <= to; i++)
            //{
            //    lstColumns.Append(new Column() { Min = (uint)i, Max = (uint)i, Width = width, CustomWidth = true });
            //}
            lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 400, CustomWidth = true });
            lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 9, CustomWidth = true });
            lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 9, CustomWidth = true });
            lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 13, CustomWidth = true });
            lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 400, CustomWidth = true });
            lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 12, CustomWidth = true });
            // Only insert the columns if we had to create a new columns element
            if (needToInsertColumns)
                Sheet.Worksheet.Append(lstColumns);

            // Get the sheetData cells
            this.SheetData = Sheet.Worksheet.GetFirstChild<SheetData>();
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
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        /// <summary>
        /// Return unique sheetId
        /// </summary>
        /// <param name="docName"></param>
        /// <returns></returns>
        public UInt32Value InsertWorksheet()
        {
            // Open the document for editing.

            // Add a blank WorksheetPart.
            var spreadSheet = spreadsheetDocument;
            WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Give the new worksheet a name.
            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            return sheetId;
        }


        //// Given a WorkbookPart, inserts a new worksheet.
        //private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        //{
        //    // Add a new worksheet part to the workbook.
        //    WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        //    newWorksheetPart.Worksheet = new Worksheet(new SheetData());
        //    newWorksheetPart.Worksheet.Save();

        //    Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        //    string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

        //    // Get a unique ID for the new sheet.
        //    uint sheetId = 1;
        //    if (sheets.Elements<Sheet>().Count() > 0)
        //    {
        //        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        //    }

        //    string sheetName = "Sheet" + sheetId;

        //    // Append the new worksheet and associate it with the workbook.
        //    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        //    sheets.Append(sheet);
        //    workbookPart.Workbook.Save();

        //    return newWorksheetPart;
        //}

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


        public static void CreateSpreadsheetWorkbook(string filepath, int sheetId, string sheetName)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
        }

    }
}
