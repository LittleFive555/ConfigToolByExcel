using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConfigToolByExcel
{
    internal static class OpenXMLHelper
    {
        public static IEnumerable<Cell> GetCellsByRow(WorksheetPart worksheetPart, int rowIndex)
        {
            return worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == rowIndex);
        }

        public static IEnumerable<Cell> GetCellsByColumn(WorksheetPart worksheetPart, string columnName)
        {
            return worksheetPart.Worksheet.Descendants<Cell>().Where(c => string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0);
        }

        public static string GetCellValue(string fileName, string sheetName, string addressName)
        {
            // 以只读权限打开电子表格
            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                // 获取工作簿部分的引用
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = document.WorkbookPart ?? document.AddWorkbookPart();

                // 找到给定名称的工作表，并使用Sheet对象用来取回第一个工作表的引用。
                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet? theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                // 如果没有找到工作表，则抛出异常
                // Throw an exception if there is no sheet.
                if (theSheet == null || theSheet.Id == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // 获取到该工作表部分的引用
                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(theSheet.Id!);


                return GetCellValue(wbPart, wsPart, addressName);
            }
        }

        public static string GetCellValue(WorkbookPart wbPart, WorksheetPart wsPart, string addressName)
        {
            // 使用它的Worksheet属性来获取与提供的单元格坐标一致的单元格
            // Use its Worksheet property to get a reference to the cell 
            // whose address matches the address you supplied.
            Cell? theCell = wsPart.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == addressName).FirstOrDefault();
            return GetCellValue(wbPart, theCell);
        }

        public static string GetCellValue(WorkbookPart wbPart, Cell theCell)
        {
            // 如果该单元格不存在，返回空字符串
            // If the cell does not exist, return an empty string.
            if (theCell == null || theCell.InnerText.Length <= 0)
            {
                return string.Empty;
            }

            string value = theCell.InnerText;

            // 如果单元格内是一个整数，结束。
            // 对于日期类型内容，下面这段代码返回该日期的序列化值。
            // 这段代码单独处理字符串和布尔类型。
            // 对于shared strings，代码在共享字符串表中寻找对应的值。
            // 对于布尔类型，代码把值转换为单词TRUE或FALSE。
            // If the cell represents an integer number, you are done. 
            // For dates, this code returns the serialized value that 
            // represents the date. The code handles strings and 
            // Booleans individually. For shared strings, the code 
            // looks up the corresponding value in the shared string 
            // table. For Booleans, the code converts the value into 
            // the words TRUE or FALSE.
            if (theCell.DataType != null)
            {
                if (theCell.DataType.Value == CellValues.SharedString)
                {
                    // 对于共享字符串的值类型，在共享字符串表中寻找值
                    // For shared strings, look up the value in the
                    // shared strings table.
                    var stringTable =
                        wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    // 如果共享字符串表丢失，则表示出错，会返回在单元格中的索引。
                    // 否则，查找共享字符串表中的正确的文本。
                    // If the shared string table is missing, something 
                    // is wrong. Return the index that is in
                    // the cell. Otherwise, look up the correct text in 
                    // the table.
                    if (stringTable != null)
                    {
                        value =
                            stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                }
                else if (theCell.DataType.Value == CellValues.Boolean)
                {
                    switch (value)
                    {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }
                }
            }

            return value;
        }

        // Given a cell name, parses the specified cell to get the column name.
        public static string GetColumnName(string? cellName)
        {
            if (cellName == null)
            {
                return string.Empty;
            }
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        // Given a cell name, parses the specified cell to get the row index.
        public static uint? GetRowIndex(string? cellName)
        {
            if (cellName == null)
            {
                return null;
            }

            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }
    }
}