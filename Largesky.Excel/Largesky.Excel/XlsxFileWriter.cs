using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Largesky.Excel
{
    public class XlsxFileWriter
    {
        /// <summary>
        /// 插入共享文字中
        /// </summary>
        /// <param name="text"></param>
        /// <param name="shareStringPart"></param>
        /// <returns></returns>
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

        public static void WriteXlsx(string file, string[][] contents)
        {
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(file, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var ssp = workbookpart.AddNewPart<SharedStringTablePart>();

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            uint rowCount = 1;
            //加入所有行
            foreach (var contentRow in contents)
            {
                Row row = new Row { RowIndex = rowCount, Spans = new ListValue<StringValue>() };
                row.Spans.Items.Add(new StringValue("1:" + contentRow.Length.ToString()));
                for (int i = 0; i < contentRow.Length; i++)
                {
                    if (contentRow[i] == null)
                    {
                        continue;
                    }
                    Cell cell = new Cell { CellReference = XlsxFileReader.GetExcelColumnName(i + 1) + rowCount };

                    if (contentRow[i].All(obj => Char.IsDigit(obj) || obj == '.'))
                    {
                        cell.CellValue = new CellValue(contentRow[i]);
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    }
                    else
                    {
                        cell.CellValue = new CellValue(InsertSharedStringItem(contentRow[i], ssp).ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }
                    row.AppendChild(cell);
                }
                sheetData.AppendChild(row);
                rowCount++;
            }
            spreadsheetDocument.Close();
        }
    }
}
