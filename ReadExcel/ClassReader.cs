using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ReadExcel
{
    internal class ClassReader
    {
        private const string OutputSymbol = "*";
        private const int PropertyOutputSymbolRowIndex = 1;
        private const int PropertyNameCellRowIndex = 2;
        private const int PropertyTypeRowIndex = 3;
        private const int CommentRowIndex = 4;
        private const int DefaultValueRowIndex = 5;
        private const int DataStartRowIndex = 6;

        private const string ValueOutputSymbolColumnName = "A";

        public static IReadOnlyList<ClassInfo>? CollectClassesInfo(string docName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
            {
                IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>();
                if (sheets == null || sheets.Count() == 0)
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

        public static Dictionary<Type, List<BaseData>>? CollectNumeric(string docName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
            {
                IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>();
                if (sheets == null || sheets.Count() == 0)
                    return null;

                Dictionary<Type, List<BaseData>> numericsByClassType = new Dictionary<Type, List<BaseData>>();
                foreach (var sheet in sheets)
                {
                    string classTypeStr = sheet.Name?.Value ?? string.Empty;
                    string classTypeWithNamespaceStr = string.Format("ReadExcel." + classTypeStr);
                    var type = Type.GetType(classTypeWithNamespaceStr);
                    if (type == null)
                    {
                        Console.WriteLine(string.Format("Warning: {0}", classTypeWithNamespaceStr));
                        continue;
                    }

                    List<BaseData>? datas = GetDatas(document, sheet, type);
                    if (datas != null && datas.Count > 0)
                        numericsByClassType.Add(type, datas);
                }
                return numericsByClassType;
            }
        }

        private static List<BaseData>? GetDatas(SpreadsheetDocument document, Sheet sheet, Type? type)
        {
            string? id = sheet.Id;
            if (id == null)
                throw new NullReferenceException(string.Format("Error: Sheet {0} has no id.", sheet.Name));

            var datas = new List<BaseData>();
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(id);
            if (type != null && type.IsSubclassOf(typeof(BaseData)))
            {
                IEnumerable<Cell> propertyNameCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == PropertyNameCellRowIndex);
                if (propertyNameCells.Count() == 0)
                    return null; // 空表

                // 构建字段名到表格列的索引
                var properties = type.GetProperties();
                Dictionary<string, string> propertyNameToColumn = new Dictionary<string, string>();
                foreach (var property in properties)
                {
                    string propertyName = property.Name;
                    var specificPropertyNameCell = propertyNameCells.First(c => GetCellText(document, c).Equals(propertyName));
                    if (specificPropertyNameCell?.CellReference == null)
                        throw new NullReferenceException(string.Format("Error: Sheet {0} has no property {1}", sheet.Name, propertyName));

                    string columnName = GetColumnName(specificPropertyNameCell.CellReference.Value);
                    propertyNameToColumn.Add(propertyName, columnName);
                }

                IEnumerable<Cell> valueOutputCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => string.Compare(GetColumnName(c.CellReference?.Value), ValueOutputSymbolColumnName, true) == 0);
                foreach (var cell in valueOutputCells)
                {
                    if (IsOutout(GetCellText(document, cell)))
                    {
                        uint rowIndex = GetRowIndex(cell.CellReference?.Value) ?? 0;
                        var objectInstance = Activator.CreateInstance(type);
                        foreach (var property in properties)
                        {
                            string propertyName = property.Name;
                            var specificPropertyValueCell = worksheetPart.Worksheet.Descendants<Cell>().First(c => string.Compare(c.CellReference?.Value, propertyNameToColumn[propertyName] + rowIndex, true) == 0);
                            string columnName = GetColumnName(specificPropertyValueCell.CellReference?.Value);
                            string valueText = GetCellText(document, specificPropertyValueCell);
                            var value = Convert.ChangeType(valueText, property.PropertyType); // TODO 是否要转换判断
                            property.SetValue(objectInstance, value);
                        }
                        datas.Add((BaseData)objectInstance);
                    }
                }
            }

            return datas;
        }

        private static ClassInfo? GetClassInfo(SpreadsheetDocument document, Sheet sheet)
        {
            ClassInfo classInfo = new ClassInfo();
            string? id = sheet.Id;
            if (id == null)
                throw new NullReferenceException(string.Format("Error: Sheet {0} has no id.", sheet.Name));

            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(id);
            IEnumerable<Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == PropertyOutputSymbolRowIndex);
            if (cells.Count() == 0) // 空表
                return null;

            // 以工作表名作为类名
            if (string.IsNullOrEmpty(sheet?.Name?.Value))
            {
                Console.WriteLine("Error: A sheet has no name.");
                return null;
            }
            classInfo.ClassName = sheet.Name.Value;

            List<PropertyInfo> propertyInfos = new List<PropertyInfo>();
            // 获取所有输出的属性名及数据类型
            foreach (var cell in cells)
            {
                if (IsOutout(GetCellText(document, cell)))
                {
                    string columnName = GetColumnName(cell.CellReference?.Value);
                    IEnumerable<Cell> nameCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == PropertyNameCellRowIndex && string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0);
                    if (nameCells.Count() == 0)
                        return null; // TODO
                    string propertyName = GetCellText(document, nameCells.First()); // TODO 对类名规范进行判断

                    IEnumerable<Cell> typeCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == PropertyTypeRowIndex && string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0);
                    if (typeCells.Count() == 0)
                        return null; // TODO
                    string propertyType = GetCellText(document, typeCells.First()); // TODO 对数据类型进行判断

                    propertyInfos.Add(new PropertyInfo() { Name = propertyName, Type = propertyType });
                }
            }
            classInfo.Properties = propertyInfos;
            return classInfo;
        }

        private static string GetCellText(SpreadsheetDocument document, Cell cell)
        {
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && int.TryParse(cell.CellValue?.Text, out int index))
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
