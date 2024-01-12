using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization.Metadata;
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

                ClassInfo baseDataInfo = new ClassInfo() 
                { 
                    ClassName = "BaseData",
                    Properties = new List<PropertyInfo>() 
                    {
                        new PropertyInfo() { Type = "int", Name = "NID" } 
                    }
                };
                classesInfo.Add(baseDataInfo);

                foreach (var sheet in sheets)
                {
                    ClassInfo? classInfo = GetClassInfo(document, sheet);
                    if (classInfo != null)
                        classesInfo.Add(classInfo.Value);
                }
                return classesInfo;
            }
        }

        public static Dictionary<string, JsonObject>? CollectNumeric(string docName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
            {
                IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>();
                if (sheets == null || sheets.Count() == 0)
                    return null;

                Dictionary<string, JsonObject> jsonNumerics = new Dictionary<string, JsonObject>();
                foreach (var sheet in sheets)
                {
                    string classTypeStr = sheet.Name?.Value ?? string.Empty;
                    JsonArray? datas = GetDatas(document, sheet);
                    if (datas != null && datas.Count > 0)
                    {
                        var jsonObject = new JsonObject
                        {
                            { "Content", datas }
                        };
                        jsonNumerics.Add(classTypeStr, jsonObject);
                    }
                }
                return jsonNumerics;
            }
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
                    if (propertyName.Equals("NID")) // 基类包含，直接跳过
                        continue;

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

        private static JsonArray? GetDatas(SpreadsheetDocument document, Sheet sheet)
        {
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

            List<PropertyInfo> propertyInfos = new List<PropertyInfo>();
            // 构建字段名到表格列的索引
            Dictionary<string, string> propertyNameToColumn = new Dictionary<string, string>();
            foreach (var cell in cells)
            {
                if (IsOutout(GetCellText(document, cell)))
                {
                    string columnName = GetColumnName(cell.CellReference?.Value);
                    IEnumerable<Cell> nameCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => (GetRowIndex(c.CellReference?.Value) ?? 0) == PropertyNameCellRowIndex && string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0);
                    if (nameCells.Count() == 0)
                        return null; // TODO
                    string propertyName = GetCellText(document, nameCells.First()); // TODO 对类名规范进行判断
                    propertyNameToColumn.Add(propertyName, columnName);
                }
            }

            JsonArray jsonArray = new JsonArray();
            IEnumerable<Cell> valueOutputCells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => string.Compare(GetColumnName(c.CellReference?.Value), ValueOutputSymbolColumnName, true) == 0);
            foreach (var cell in valueOutputCells)
            {
                if (IsOutout(GetCellText(document, cell)))
                {
                    JsonObject jsonObject = new JsonObject();
                    foreach (var propertyName in propertyNameToColumn.Keys)
                    {
                        string columnName = propertyNameToColumn[propertyName];
                        uint valueRowIndex = GetRowIndex(cell.CellReference?.Value) ?? 0;
                        var specificPropertyValueCell = worksheetPart.Worksheet.Descendants<Cell>().First(c => string.Compare(c.CellReference?.Value, columnName + valueRowIndex, true) == 0);
                        string valueText = GetCellText(document, specificPropertyValueCell);

                        var specificPropertyValueTypeCell = worksheetPart.Worksheet.Descendants<Cell>().First(c => string.Compare(c.CellReference?.Value, columnName + PropertyTypeRowIndex, true) == 0);
                        string propertyValueTypeText = GetCellText(document, specificPropertyValueTypeCell);

                        string fullTypeName = GetFullTypeName(propertyValueTypeText);
                        Type? type = Type.GetType(fullTypeName);
                        if (type == null)
                        {
                            Console.WriteLine($"Error: Didn't find type <{propertyValueTypeText}>.");
                            return null;
                        }
                        var value = Convert.ChangeType(valueText, type); // TODO 是否要转换判断
                        var jsonNode = JsonSerializer.SerializeToNode(value, JsonTypeInfo.CreateJsonTypeInfo(type, JsonSerializerOptions.Default));
                        jsonObject.Add(propertyName, jsonNode);
                    }
                    jsonArray.Add(jsonObject);
                }
            }
            return jsonArray;
        }

        private static string GetFullTypeName(string input)
        {
            switch(input)
            {
                case "int":
                    return "System.Int32";
                case "float":
                    return "System.Single";
                case "string":
                    return "System.String";
            }
            return string.Empty;
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
