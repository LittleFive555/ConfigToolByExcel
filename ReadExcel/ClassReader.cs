using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ReadExcel
{
    internal class ClassReader
    {
        private const string OutputSymbol = "*";

        public static IReadOnlyList<ClassInfo>? CollectClassesInfo(string docName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
            {
                IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>();
                if (sheets is null || sheets.Count() == 0)
                    return null;

                List<ClassInfo> classesInfo = new List<ClassInfo>();
                foreach (var sheet in sheets)
                {
                    ClassInfo? classInfo = GetClassInfo(document, sheet);
                    if (classInfo != null)
                        classesInfo.Add(classInfo.Value);
                }
                return classesInfo;
            }
        }

        private static ClassInfo? GetClassInfo(SpreadsheetDocument document, Sheet sheet)
        {
            ClassInfo classInfo = new ClassInfo();
            string? id = sheet.Id;
            if (id is null)
                return null;
            int outputSymbolRowIndex = 1;
            int fieldNameCellRowIndex = 2;
            int fieldTypeRowIndex = 3;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(id);
            IEnumerable<Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == outputSymbolRowIndex);
            if (cells.Count() == 0)
                return null; // TODO

            // 以工作表名作为类名
            if (string.IsNullOrEmpty(sheet?.Name?.Value))
                return null; // TODO
            classInfo.ClassName = sheet.Name.Value;

            List<FieldInfo> fieldInfos = new List<FieldInfo>();
            // 获取所有输出的属性名及数据类型
            foreach (var cell in cells)
            {
                if (IsOutout(GetCellText(document, cell)))
                {
                    string columnName = GetColumnName(cell.CellReference?.Value);
                    IEnumerable<Cell> nameCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == fieldNameCellRowIndex && string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0);
                    if (nameCells.Count() == 0)
                        return null; // TODO
                    string fieldName = GetCellText(document, nameCells.First()); // TODO 对类名规范进行判断

                    IEnumerable<Cell> typeCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == fieldTypeRowIndex && string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0);
                    if (typeCells.Count() == 0)
                        return null; // TODO
                    string fieldType = GetCellText(document, typeCells.First()); // TODO 对数据类型进行判断

                    fieldInfos.Add(new FieldInfo() { Name = fieldName, Type = fieldType });
                }
            }
            classInfo.Fields = fieldInfos;
            return classInfo;
        }

        private static string GetCellText(SpreadsheetDocument document, Cell cell)
        {
            if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString && int.TryParse(cell.CellValue?.Text, out int index))
            {
                SharedStringTablePart shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                SharedStringItem[] items = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();

                return items[index].InnerText;
            }
            else
            {
                return cell.CellValue?.Text;
            }
        }

        private static bool IsOutout(string? cellValue)
        {
            if (cellValue is null)
                return false;

            return cellValue.Equals(OutputSymbol);
        }

        // Given a cell name, parses the specified cell to get the column name.
        private static string GetColumnName(string? cellName)
        {
            if (cellName is null)
            {
                return string.Empty;
            }
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        // Given a cell name, parses the specified cell to get the row index.
        private static uint? GetRowIndex(string? cellName)
        {
            if (cellName is null)
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
