using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Largesky.Excel
{
    public class XlsxFileReader
    {
        private static readonly char[] NUMBERS = new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
        private Dictionary<string, string[][]> sheetDatas = new Dictionary<string, string[][]>();

        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            if (columnNumber < 1)
            {
                throw new ArgumentException("columnNumber 不能小于1");
            }

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static int GetExcelColumnIndex(string colName)
        {
            if (string.IsNullOrWhiteSpace(colName))
            {
                throw new ArgumentException("colName不能为空");
            }

            var colIndex = 0;
            for (int ind = 0, pow = colName.Count() - 1; ind < colName.Count(); ++ind, --pow)
            {
                var cVal = Convert.ToInt32(colName[ind]) - 64; //col A is index 1
                colIndex += cVal * ((int)Math.Pow(26, pow));
            }
            return colIndex;
        }

        private static string ReadCellValue(SpreadsheetDocument doc, Worksheet sheet, Cell cell)
        {
            string val = null;
            SharedStringTablePart sharedStringPart = doc.WorkbookPart.SharedStringTablePart;

            if (cell.CellValue == null)
            {
                return null;
            }

            if (cell.DataType == null)
            {
                val = cell.CellValue.InnerText;
            }
            else if (cell.DataType == CellValues.Date)
            {
                throw new NotImplementedException(" ReadCellValue Date");
            }
            else if (cell.DataType == CellValues.Boolean)
            {
                throw new NotImplementedException(" ReadCellValue Date");
            }
            else if (cell.DataType == CellValues.Error)
            {
                throw new NotImplementedException(" ReadCellValue Date");
            }
            else if (cell.DataType == CellValues.InlineString)
            {
                val = cell.CellValue.InnerText;
            }
            else if (cell.DataType == CellValues.Number)
            {
                throw new NotImplementedException(" ReadCellValue Date");
            }
            else if (cell.DataType == CellValues.SharedString)
            {
                int strIndex = int.Parse(cell.CellValue.InnerText);
                var item = sharedStringPart.SharedStringTable.ElementAt(strIndex).FirstChild;
                val = item.InnerText;
            }
            else if (cell.DataType == CellValues.String)
            {
                return cell.CellValue.InnerText;
            }
            else
            {
                val = cell.CellValue.InnerText;
            }
            return val;
        }

        private static string[][] ParseSheet(SpreadsheetDocument doc, Worksheet sheet, string sheetName)
        {
            SheetData sd = sheet.OfType<SheetData>().FirstOrDefault();
            if (sd == null)
            {
                return new string[0][] { };
            }
            List<string[]> datas = new List<string[]>();
            var rows = sd.OfType<Row>().Where(obj => obj.Spans != null && obj.Spans.Items != null && obj.Spans.Items.First() != null).ToArray();
            //检查所有行中最大的列
            var rowSpans = rows.Select(obj => obj.Spans.Items.First().Value).ToArray();
            rowSpans = rowSpans.Select(obj => obj.Substring(obj.IndexOf(':') + 1)).ToArray();
            int maxCount = rowSpans.Select(obj => int.Parse(obj)).Max();
            foreach (var row in rows)
            {
                var data = new string[maxCount + 1];
                data[0] = row.RowIndex.Value.ToString();
                foreach (var cell in row.OfType<Cell>())
                {
                    try
                    {
                        var v = cell.CellValue;
                        string address = cell.CellReference;//行列位置如 A1,B1
                        int index = address.IndexOfAny(NUMBERS);
                        if (index < 1)
                        {
                            throw new Exception("表格地址不对无法解析");
                        }
                        int colNumber = GetExcelColumnIndex(address.Substring(0, index));
                        data[colNumber] = ReadCellValue(doc, sheet, cell);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(string.Format("解析表:{0},位置:{1}{2}出错{3}", sheetName, cell.CellReference, Environment.NewLine, ex.Message), ex);
                    }
                }
                datas.Add(data);
            }
            return datas.ToArray();
        }

        /// <summary>
        /// 打开一个文件
        /// </summary>
        /// <param name="file">要打开的文件路径</param>
        /// <returns></returns>
        public static XlsxFileReader Open(string file, string sheetName = "Sheet1")
        {
            XlsxFileReader fr = new XlsxFileReader();
            SpreadsheetDocument tmpDoc = null;

            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.ReadWrite))
            {
                tmpDoc = SpreadsheetDocument.Open(fs, true);
                var workbookPart = tmpDoc.WorkbookPart;
                var workBook = workbookPart.Workbook;

                if (string.IsNullOrWhiteSpace(sheetName) == false)
                {
                    if (workbookPart.Workbook.Sheets.FirstOrDefault(obj => (obj as Sheet).Name == sheetName) == null)
                    {
                        throw new Exception("指定的文件中不包含:" + sheetName + "文件:" + file);
                    }
                }
                foreach (Sheet sheet in tmpDoc.WorkbookPart.Workbook.Sheets)
                {
                    var vvvvvvvvvvv = tmpDoc.WorkbookPart.GetPartById(sheet.Id);
                    var s = tmpDoc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;
                    if (s == null)
                    {
                        throw new Exception("文件结构不对，找不表:" + sheet.Name + "的数据" + "文件:" + file);
                    }
                    //为空表示需要解析所有的表
                    if (string.IsNullOrWhiteSpace(sheetName) || sheetName.Equals(sheet.Name))
                    {
                        string[][] rows = ParseSheet(tmpDoc, s.Worksheet, sheet.Name);
                        fr.sheetDatas.Add(sheet.Name, rows);
                    }
                }
            }

            return fr;
        }

        public string[][] ReadAllRows(string sheetName = "Sheet1")
        {
            if (sheetDatas.ContainsKey(sheetName) == false)
            {
                throw new Exception("指定的表不存在");
            }
            return sheetDatas[sheetName];
        }
    }
}
