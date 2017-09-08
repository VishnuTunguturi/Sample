﻿/*using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Exportc
{
    class MainClass
    {
        public static void Main(string[] args)
        {
			ExportDataSet("test.xlsx");
			Console.WriteLine("File Export successfully...");
			Console.ReadLine();
        }

		private static void ExportDataSet(string destination)
		{
			using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
			{
				var workbookPart = workbook.AddWorkbookPart();

				workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

				workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();



				var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
				var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
				sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

				DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
				string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

				uint sheetId = 1;
				if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
				{
					sheetId =
						sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
				}

				DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = "Test" };
				sheets.Append(sheet);

				DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

				List<String> columns = new List<string>();
				for (int i = 0; i < 5; i++)
				{
					columns.Add("column " + i);

					DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
					cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
					cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("column" + i);
					headerRow.AppendChild(cell);
				}


				sheetData.AppendChild(headerRow);

				for (int i = 0; i < 10; i++)
				{
					DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
					foreach (String col in columns)
					{
						DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
						cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
						cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Test" + i); //
						newRow.AppendChild(cell);
					}

					sheetData.AppendChild(newRow);
				}


			}
		}
    }
}*/
