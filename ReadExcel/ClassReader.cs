using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ReadExcel
{
    internal class ClassReader
    {
        private const string OutputSymbol = "*";
        private const int FieldOutputSymbolRowIndex = 1;
        private const int FieldNameCellRowIndex = 2;
        private const int FieldTypeRowIndex = 3;
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

        public static Dictionary<Type, List<BaseData>> CollectNumeric(string docName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
            {
                IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>();
                if (sheets == null || sheets.Count() == 0)
                    return null; // TODO

                Dictionary<Type, List<BaseData>> numericsByClassType = new Dictionary<Type, List<BaseData>>();
                foreach (var sheet in sheets)
                {
                    string? id = sheet.Id;
                    if (id == null)
                        return null; // TODO

                    WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(id);
                    string classTypeStr = sheet.Name.Value;
                    var type = Type.GetType(string.Format("ReadExcel." + classTypeStr));
                    if (type.IsSubclassOf(typeof(BaseData)))
                    {
                        var properties = type.GetProperties();
                        IEnumerable<Cell> propertyNameCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == FieldNameCellRowIndex);
                        if (propertyNameCells.Count() == 0)
                            return null; // TODO

                        // 构建字段名到表格列的索引
                        Dictionary<string, string> propertyNameToColumn = new Dictionary<string, string>();
                        foreach (var property in properties)
                        {
                            string propertyName = property.Name;
                            var specificPropertyNameCell = propertyNameCells.Where(c => GetCellText(document, c).Equals(propertyName));
                            if (specificPropertyNameCell.Count() == 0)
                                return null; //  TODO
                            string columnName = GetColumnName(specificPropertyNameCell.First().CellReference.Value);
                            propertyNameToColumn.Add(propertyName, columnName);
                        }

                        IEnumerable<Cell> valueOutputCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => string.Compare(GetColumnName(c.CellReference?.Value), ValueOutputSymbolColumnName, true) == 0);
                        foreach (var cell in valueOutputCells)
                        {
                            if (IsOutout(GetCellText(document, cell)))
                            {
                                uint rowIndex = GetRowIndex(cell.CellReference.Value) ?? 0;
                                var objectInstance = Activator.CreateInstance(type);
                                foreach (var property in properties)
                                {
                                    string propertyName = property.Name;
                                    var specificPropertyValueCell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => string.Compare(c.CellReference?.Value, propertyNameToColumn[propertyName] + rowIndex, true) == 0);
                                    if (specificPropertyValueCell.Count() == 0)
                                        return null; //  TODO
                                    string columnName = GetColumnName(specificPropertyValueCell.First().CellReference.Value);
                                    string valueText = GetCellText(document, specificPropertyValueCell.First());
                                    var value = Convert.ChangeType(valueText, property.PropertyType); // TODO 是否要转换判断
                                    property.SetValue(objectInstance, value);
                                }

                                if (numericsByClassType.ContainsKey(type))
                                    numericsByClassType[type].Add((BaseData)objectInstance);
                                else
                                    numericsByClassType.Add(type, new List<BaseData>() { (BaseData)objectInstance });
                            }
                        }
                    }
                }
                return numericsByClassType;
            }
        }

        private static ClassInfo? GetClassInfo(SpreadsheetDocument document, Sheet sheet)
        {
            ClassInfo classInfo = new ClassInfo();
            string? id = sheet.Id;
            if (id is null)
                return null;

            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(id);
            IEnumerable<Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == FieldOutputSymbolRowIndex);
            if (cells.Count() == 0)
                return null; // TODO

            // 以工作表名作为类名
            if (string.IsNullOrEmpty(sheet?.Name?.Value))
                return null; // TODO
            classInfo.ClassName = sheet.Name.Value;

            List<PropertyInfo> propertyInfos = new List<PropertyInfo>();
            // 获取所有输出的属性名及数据类型
            foreach (var cell in cells)
            {
                if (IsOutout(GetCellText(document, cell)))
                {
                    string columnName = GetColumnName(cell.CellReference?.Value);
                    IEnumerable<Cell> nameCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == FieldNameCellRowIndex && string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0);
                    if (nameCells.Count() == 0)
                        return null; // TODO
                    string propertyName = GetCellText(document, nameCells.First()); // TODO 对类名规范进行判断

                    IEnumerable<Cell> typeCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == FieldTypeRowIndex && string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0);
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
