using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Pathway.WPF.ImportExport.Logs.Domain;

namespace Pathway.WPF.ImportExport.Excel
{
	public class BaseOpenXML
    {
		protected const uint DEFAULTINDEXSTYLE = 1;
		protected const uint BOLDINDEXSTYLE = 2;
		protected const uint HEADER1INDEXSTYLE = 3;
		protected string fileName;

		protected SpreadsheetDocument SpreadsheetDocument;
		protected WorkbookPart book;

		public BaseOpenXML(string fileName)
        {
			if (String.IsNullOrEmpty(fileName))
				throw new ArgumentException("fileName");
			this.fileName = fileName;

			SpreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
			book = SpreadsheetDocument.AddWorkbookPart();
			book.Workbook = new Workbook();
			WorkbookStylesPart stylesPart = SpreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
			//Default = 1; Bold = 2; Header1 = 3
			CreatingStyle(stylesPart);
		}

		public BaseOpenXML(string fileName, bool withReopen)
		{
			if (string.IsNullOrEmpty(fileName))
			{
				throw new ArgumentException("fileName");
			}

			this.fileName = fileName;
			SpreadsheetDocument = SpreadsheetDocument.Open(fileName, isEditable: true);
			book = SpreadsheetDocument.WorkbookPart;
		}

		private void CreatingStyle(WorkbookStylesPart stylesPart)
		{
			stylesPart.Stylesheet = new Stylesheet();
			stylesPart.Stylesheet.Fonts = new Fonts();
			stylesPart.Stylesheet.Fonts.Count = 2;
			stylesPart.Stylesheet.Fonts.AppendChild(new Font()
			{
				Bold = new Bold() { Val = BooleanValue.FromBoolean(false) },
				FontName = new FontName() { Val = "Calibri" },
				FontSize = new FontSize() { Val = 11 }
			});
			stylesPart.Stylesheet.Fonts.AppendChild(new Font()
			{
				Bold = new Bold() { Val = BooleanValue.FromBoolean(true) },
				FontName = new FontName() { Val = "Arial" },
				FontSize = new FontSize() { Val = 10 }
			});
			stylesPart.Stylesheet.Fonts.Count = 2;

			// create fills
			stylesPart.Stylesheet.Fills = new Fills();

			var whiteColor = new PatternFill() { PatternType = PatternValues.Solid };
			whiteColor.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFFFFFF") };
			whiteColor.BackgroundColor = new BackgroundColor { Indexed = 64 };

			var grayColor = new PatternFill() { PatternType = PatternValues.Solid };
			grayColor.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FF808080") };
			grayColor.BackgroundColor = new BackgroundColor { Indexed = 64 };

			stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
			stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
			stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = whiteColor });
			stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = grayColor });
			stylesPart.Stylesheet.Fills.Count = 4;

			// blank border list
			stylesPart.Stylesheet.Borders = new Borders();
			stylesPart.Stylesheet.Borders.AppendChild(new Border());
			stylesPart.Stylesheet.Borders.Count = 1;

			// blank cell format list
			stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
			stylesPart.Stylesheet.CellStyleFormats.Count = 1;
			stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

			stylesPart.Stylesheet.CellFormats = new CellFormats();
			stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
			stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 0, BorderId = 0, FillId = 0, ApplyFill = true }).AppendChild(new Alignment() { WrapText = true });
			stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 0, ApplyFill = true }).AppendChild(new Alignment() { WrapText = true });
			stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 3, ApplyFill = true }).AppendChild(new Alignment() { WrapText = true });
			stylesPart.Stylesheet.CellFormats.Count = 4;

			stylesPart.Stylesheet.Save();
		}

		/// </summary>
		/// <param name="row">Row</param>
		/// <param name="value">string</param>
		/// <param name="dateType">CellValues : Default CellValues.String</param>
		/// <param name="styleIndex">styleIndex : Default = 1; Bold = 2; Header1 = 3</param>
		public void InsertCell(Row row, string value, CellValues dateType = CellValues.String, uint styleIndex = DEFAULTINDEXSTYLE)
		{
			row.AppendChild(new Cell() { CellValue = new CellValue(value), DataType = dateType, StyleIndex = styleIndex });
		}

		public void InsertColumn(Row row, string value, CellValues dateType = CellValues.String, uint styleIndex = DEFAULTINDEXSTYLE)
		{
			var c = new Column() { };
			row.AppendChild(new Cell() { CellValue = new CellValue(value), DataType = dateType, StyleIndex = styleIndex });
		}

		protected Columns SetDefaultColumnsWidth(double defaultColumnWidth)
		{
			Columns columns = new Columns();
			Column column = new Column() { Min = 1, Max = 255, Width = defaultColumnWidth, CustomWidth = true };
			columns.Append(column);
			return columns;
		}

		/// <summary>
		/// Saves data
		/// </summary>
		public void SaveAndClose()
		{
			this.book.Workbook.Save();
			this.SpreadsheetDocument.Close();
		}

		/// <summary>
		/// Create sheet
		/// </summary>
		/// <returns>WorksheetPart and SheetData</returns>
		protected KeyValuePair<WorksheetPart, SheetData> CreateSheet(string sheetName, uint shitId, double defaultColumnWidth = 19)
		{
			WorksheetPart wsPart = SpreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
			wsPart.Worksheet = new Worksheet();
			wsPart.Worksheet.Append(SetDefaultColumnsWidth(defaultColumnWidth));

			SheetData sheetData = wsPart.Worksheet.AppendChild(new SheetData());

			wsPart.Worksheet.Save();

			Sheets sheets = SpreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
			if (sheets == null)
				sheets = SpreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
			sheets.AppendChild(new Sheet() { Id = SpreadsheetDocument.WorkbookPart.GetIdOfPart(wsPart), SheetId = shitId, Name = sheetName });

			return new KeyValuePair<WorksheetPart, SheetData>(wsPart, sheetData);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="defaultColumnWidth"></param>
		/// <returns></returns>
		protected virtual Columns SetColumnsWidth(IList<ColumnsPreference> columnsPreferences)
		{
			Columns columns = new Columns();
			if (columnsPreferences == null)
			{
				Column column = new Column() { Min = 1, Max = 255, Width = 50, CustomWidth = true };
				columns.Append(column);
			}
			else
			{
				foreach (ColumnsPreference p in columnsPreferences)
				{
					Column column = new Column() { Min = p.Min, Max = p.Max, Width = p.Width, CustomWidth = true };
					columns.Append(column);
				}
			}

			return columns;
		}
	}
}
