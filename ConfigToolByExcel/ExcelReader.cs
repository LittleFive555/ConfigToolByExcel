using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using ConfigToolByExcel.CodeGenerator;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConfigToolByExcel
{
    internal class ExcelReader
    {
        private const string OutputSymbol = "*";
        private const int PropertyOutputSymbolRowIndex = 1;
        private const int PropertyNameCellRowIndex = 2;
        private const int PropertyTypeRowIndex = 3;
        private const int CommentRowIndex = 4;
        private const int DefaultValueRowIndex = 5;
        private const int DataStartRowIndex = 6;

        private const string ValueOutputSymbolColumnName = "A";

        private const string JsonObjectName = "Content";

        private const string IDColumnName = "ID";

        private static readonly Regex ClassNameRegex = new Regex("^[A-Z][A-Za-z0-9_]*");
        private static readonly Regex PropertyNameRegex = new Regex("^[A-Z][A-Za-z0-9_]*");

        public static IReadOnlyList<TableInfo>? CollectTableInfo(string docName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
            {
                IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>();
                if (sheets == null || sheets.Count() == 0)
                    return null;

                List<TableInfo> tablesInfo = new List<TableInfo>();

                // 从配置表中获取自定义的配置类信息，一个工作表代表一个类
                foreach (var sheet in sheets)
                {
                    TableInfo? tableInfo = GetTableInfo(document, sheet);
                    if (tableInfo != null)
                        tablesInfo.Add(tableInfo.Value);
                }

                return tablesInfo;
            }
        }

        public static Dictionary<string, JsonObject>? CollectData(string docName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
            {
                IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>();
                if (sheets == null || sheets.Count() == 0)
                    return null;

                // 一个工作簿中可能包含多个工作表，一个工作表为一个配置类
                // 字典的key为类名（与工作表名相同），value为数据转换后的JsonObject对象
                Dictionary<string, JsonObject> jsonDatas = new Dictionary<string, JsonObject>();

                foreach (var sheet in sheets)
                {
                    string classTypeStr = sheet.Name?.Value ?? string.Empty;

                    // 从工作表中获取配置数据
                    JsonArray? datas = GetDatas(document, sheet);
                    if (datas != null && datas.Count > 0)
                    {
                        var jsonObject = new JsonObject { { JsonObjectName, datas } };
                        jsonDatas.Add(classTypeStr, jsonObject);
                    }
                }

                return jsonDatas;
            }
        }

        private static TableInfo? GetTableInfo(SpreadsheetDocument document, Sheet sheet)
        {
            TableInfo tableInfo = new TableInfo();
            string? id = sheet.Id;
            if (id == null)
                throw new NullReferenceException(string.Format("Error: Sheet {0} has no id.", sheet.Name));

            WorkbookPart wbPart = document.WorkbookPart ?? document.AddWorkbookPart();
            WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(id);
            IEnumerable<Cell> propertyOutputSymbolCells = OpenXMLHelper.GetCellsByRow(worksheetPart, PropertyOutputSymbolRowIndex);
            if (propertyOutputSymbolCells.Count() == 0) // 空表
                return null;

            // 以工作表名作为类名
            if (string.IsNullOrEmpty(sheet?.Name?.Value))
            {
                Console.WriteLine("Error: A sheet has no name.");
                return null;
            }
            // 对类名规范进行判断
            if (!IsValidClassName(sheet.Name.Value))
                throw new FormatException($"Invalid class name <{sheet.Name.Value}>. Regex pattern {ClassNameRegex}");
            tableInfo.TableName = sheet.Name.Value;

            string dataIdType = string.Empty;

            List<FieldInfo> propertyInfos = new List<FieldInfo>();
            // 获取所有输出的属性名及数据类型
            foreach (var outputSymbolCell in propertyOutputSymbolCells)
            {
                if (IsOutput(OpenXMLHelper.GetCellValue(wbPart, outputSymbolCell)))
                {
                    string columnName = OpenXMLHelper.GetColumnName(outputSymbolCell.CellReference?.Value);

                    string propertyName = OpenXMLHelper.GetCellValue(wbPart, worksheetPart, columnName + PropertyNameCellRowIndex);
                    // 对属性名规范进行判断
                    if (!IsValidPropertyName(propertyName))
                        throw new FormatException($"Invalid property name <{propertyName}> in table <{tableInfo.TableName}>. Regex pattern {PropertyNameRegex}");

                    string propertyType = OpenXMLHelper.GetCellValue(wbPart, worksheetPart, columnName + PropertyTypeRowIndex);
                    // 对属性类型规范进行判断
                    if (!ValueConverter.IsValidType(propertyType))
                        throw new InvalidCastException($"Invalid property type <{propertyType}> for property <{propertyName}> in table <{tableInfo.TableName}>.");

                    if (propertyName == IDColumnName)
                        dataIdType = propertyType;
                    propertyInfos.Add(new FieldInfo() { Name = propertyName, Type = propertyType });
                }
            }

            if (propertyInfos.Count <= 0) // 没有输出的属性，为空类
                return null;

            if (string.IsNullOrEmpty(dataIdType))
                throw new FileFormatException($"Table <{tableInfo.TableName}> must have a property named <{IDColumnName}> to represent the data ID.");

            tableInfo.IDType = dataIdType;
            tableInfo.Fields = propertyInfos;
            return tableInfo;
        }

        private static JsonArray? GetDatas(SpreadsheetDocument document, Sheet sheet)
        {
            string? id = sheet.Id;
            if (id == null)
                throw new NullReferenceException(string.Format("Error: Sheet {0} has no id.", sheet.Name));

            WorkbookPart wbPart = document.WorkbookPart ?? document.AddWorkbookPart();
            WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(id);

            IEnumerable<Cell> propertyOutputSymbolCells = OpenXMLHelper.GetCellsByRow(worksheetPart, PropertyOutputSymbolRowIndex);
            if (propertyOutputSymbolCells.Count() == 0) // 空表
                return null;

            // 以工作表名作为类名
            if (string.IsNullOrEmpty(sheet?.Name?.Value))
            {
                Console.WriteLine("Error: A sheet has no name.");
                return null;
            }
            // 对类名规范进行判断
            if (!IsValidClassName(sheet.Name.Value))
                throw new FormatException($"Invalid class name <{sheet.Name.Value}>. Regex pattern {ClassNameRegex}");

            // 构建字段名到表格列的索引
            Dictionary<string, string> propertyNameToColumn = new Dictionary<string, string>();
            foreach (var outputSymbolCell in propertyOutputSymbolCells)
            {
                if (IsOutput(OpenXMLHelper.GetCellValue(wbPart, outputSymbolCell)))
                {
                    string columnName = OpenXMLHelper.GetColumnName(outputSymbolCell.CellReference?.Value);
                    string propertyName = OpenXMLHelper.GetCellValue(wbPart, worksheetPart, columnName + PropertyNameCellRowIndex);
                    // 对属性名规范进行判断
                    if (!IsValidPropertyName(propertyName))
                        throw new FormatException($"Invalid property name <{propertyName}> in class <{sheet.Name.Value}>. Regex pattern {PropertyNameRegex}.");
                    propertyNameToColumn.Add(propertyName, columnName);
                }
            }

            JsonArray jsonArray = new JsonArray();
            IEnumerable<Cell> valueOutputCells = OpenXMLHelper.GetCellsByColumn(worksheetPart, ValueOutputSymbolColumnName);
            foreach (var cell in valueOutputCells)
            {
                // 如果该行数据需要输出，才读取数据并转换json
                if (IsOutput(OpenXMLHelper.GetCellValue(wbPart, cell)))
                {
                    JsonObject jsonObject = new JsonObject();
                    uint valueRowIndex = OpenXMLHelper.GetRowIndex(cell.CellReference?.Value) ?? 0;
                    foreach (var propertyName in propertyNameToColumn.Keys)
                    {
                        string columnName = propertyNameToColumn[propertyName];
                        string propertyValueTypeText = OpenXMLHelper.GetCellValue(wbPart, worksheetPart, columnName + PropertyTypeRowIndex);
                        string valueText = OpenXMLHelper.GetCellValue(wbPart, worksheetPart, columnName + valueRowIndex);
                        if (string.IsNullOrEmpty(valueText)) // 对应格子未配备时，使用默认值
                            valueText = OpenXMLHelper.GetCellValue(wbPart, worksheetPart, columnName + DefaultValueRowIndex);
                        // 获取数据类型和值
                        if (!ValueConverter.TryConvertValue(propertyValueTypeText, valueText, out Type? type, out object? value))
                            throw new InvalidCastException($"Invalid value type <{propertyValueTypeText}>.");

                        // 转换为JsonNode
                        var jsonNode = JsonSerializer.SerializeToNode(value, type);
                        jsonObject.Add(propertyName, jsonNode);
                    }
                    jsonArray.Add(jsonObject);
                }
            }
            return jsonArray;
        }

        private static bool IsOutput(string? cellValue)
        {
            if (cellValue is null)
                return false;

            return cellValue.Equals(OutputSymbol);
        }

        private static bool IsValidClassName(string str)
        {
            return ClassNameRegex.IsMatch(str);
        }

        private static bool IsValidPropertyName(string str)
        {
            return PropertyNameRegex.IsMatch(str);
        }
    }
}
