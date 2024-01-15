using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization.Metadata;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConfigToolByExcel
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

        private const string JsonObjectName = "Content";

        public static IReadOnlyList<ClassInfo>? CollectClassesInfo(string docName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
            {
                IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>();
                if (sheets == null || sheets.Count() == 0)
                    return null;

                List<ClassInfo> classesInfo = new List<ClassInfo>();

                // 添加基类信息
                ClassInfo baseDataInfo = new ClassInfo() 
                { 
                    ClassName = "BaseData",
                    Properties = new List<PropertyInfo>() 
                    {
                        new PropertyInfo() { Type = "int", Name = "NID" } 
                    }
                };
                classesInfo.Add(baseDataInfo);

                // 从配置表中获取自定义的配置类信息，一个工作表代表一个类
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

                // 一个工作簿中可能包含多个工作表，一个工作表为一个配置类
                // 字典的key为类名（与工作表名相同），value为数据转换后的JsonObject对象
                Dictionary<string, JsonObject> jsonNumerics = new Dictionary<string, JsonObject>();

                foreach (var sheet in sheets)
                {
                    string classTypeStr = sheet.Name?.Value ?? string.Empty;

                    // 从工作表中获取配置数据
                    JsonArray? datas = GetDatas(document, sheet);

                    // 将配置数据转换为json格式
                    if (datas != null && datas.Count > 0)
                    {
                        var jsonObject = new JsonObject{ { JsonObjectName, datas } };
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
            // TODO 对类名规范进行判断
            classInfo.ClassName = sheet.Name.Value;

            List<PropertyInfo> propertyInfos = new List<PropertyInfo>();
            // 获取所有输出的属性名及数据类型
            foreach (var outputSymbolCell in propertyOutputSymbolCells)
            {
                if (IsOutput(OpenXMLHelper.GetCellValue(wbPart, outputSymbolCell)))
                {
                    string columnName = OpenXMLHelper.GetColumnName(outputSymbolCell.CellReference?.Value);
                    // TODO 对属性名规范进行判断
                    string propertyName = OpenXMLHelper.GetCellValue(wbPart, worksheetPart, columnName + PropertyNameCellRowIndex);
                    if (propertyName.Equals("NID")) // 基类包含，直接跳过
                        continue;
                    // TODO 对类型进行判断
                    string propertyType = OpenXMLHelper.GetCellValue(wbPart, worksheetPart, columnName + PropertyTypeRowIndex);
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

            WorkbookPart wbPart = document.WorkbookPart ?? document.AddWorkbookPart();
            WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(id);

            IEnumerable<Cell> propertyOutputSymbolCells = OpenXMLHelper.GetCellsByRow(worksheetPart, PropertyOutputSymbolRowIndex);
            if (propertyOutputSymbolCells.Count() == 0) // 空表
                return null;

            // 构建字段名到表格列的索引
            Dictionary<string, string> propertyNameToColumn = new Dictionary<string, string>();
            foreach (var outputSymbolCell in propertyOutputSymbolCells)
            {
                if (IsOutput(OpenXMLHelper.GetCellValue(wbPart, outputSymbolCell)))
                {
                    string columnName = OpenXMLHelper.GetColumnName(outputSymbolCell.CellReference?.Value);
                    // TODO 对属性名规范进行判断
                    string propertyName = OpenXMLHelper.GetCellValue(wbPart, worksheetPart, columnName + PropertyNameCellRowIndex);
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
                    foreach (var propertyName in propertyNameToColumn.Keys)
                    {
                        string columnName = propertyNameToColumn[propertyName];

                        // 获取数据类型
                        string propertyValueTypeText = OpenXMLHelper.GetCellValue(wbPart, worksheetPart, columnName + PropertyTypeRowIndex);
                        string fullTypeName = GetFullTypeName(propertyValueTypeText);
                        Type? type = Type.GetType(fullTypeName);
                        if (type == null)
                        {
                            Console.WriteLine($"Error: Didn't find type <{propertyValueTypeText}>.");
                            return null;
                        }

                        // 获取数据并转换类型
                        uint valueRowIndex = OpenXMLHelper.GetRowIndex(cell.CellReference?.Value) ?? 0;
                        string valueText = OpenXMLHelper.GetCellValue(wbPart, worksheetPart, columnName + valueRowIndex);
                        var value = Convert.ChangeType(valueText, type); // TODO 是否要转换判断

                        // 转换为JsonNode
                        var jsonNode = JsonSerializer.SerializeToNode(value, type);
                        
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

        private static bool IsOutput(string? cellValue)
        {
            if (cellValue is null)
                return false;

            return cellValue.Equals(OutputSymbol);
        }
    }
}
