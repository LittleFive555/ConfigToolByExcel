namespace ConfigToolByExcel.CodeGenerator
{
    internal class GoCodeGenerator : BaseCodeGenerator
    {
        public static void GenerateGoFile(string packageName, TableInfo tableInfo, string outputPath)
        {
            string fileName = string.Format("{0}.go", tableInfo.TableName.ToLower());
            string fullPath = Path.Combine(outputPath, fileName);
            // 如果文件存在，先删除
            if (File.Exists(fullPath))
                File.Delete(fullPath);
            using (FileStream fileStream = File.Create(fullPath))
            {
                int level = 0;
                AddLine(fileStream, level, $"package {packageName}");
                AddLine(fileStream, level, string.Empty);
                // struct start
                AddLine(fileStream, level, $"type D{tableInfo.TableName} struct {{");
                level++;
                // property start
                if (tableInfo.Fields != null)
                {
                    foreach (FieldInfo fieldInfo in tableInfo.Fields)
                        AddLine(fileStream, level, $"{fieldInfo.Name} {fieldInfo.Type}");
                }
                // property end
                level--;
                AddLine(fileStream, level, "}");
                AddLine(fileStream, level, string.Empty);

                AddLine(fileStream, level, $"func (data D{tableInfo.TableName}) GetID() {tableInfo.IDType} {{"); // TODO 类型转换为golang类型
                level++;
                AddLine(fileStream, level, $"return data.ID");
                level--;
                AddLine(fileStream, level, "}");

                AddLine(fileStream, level, $"type D{tableInfo.TableName}List struct {{");
                level++;
                AddLine(fileStream, level, $"Content []D{tableInfo.TableName}");
                level--;
                AddLine(fileStream, level, "}");
            }
        }

        public static void GenerateInterfaceFile(string packageName, string outputPath)
        {
            string fullPath = Path.Combine(outputPath, "basedata.go");
            // 如果文件存在，先删除
            if (File.Exists(fullPath))
                File.Delete(fullPath);
            using (FileStream fileStream = File.Create(fullPath))
            {
                int level = 0;
                AddLine(fileStream, level, $"package {packageName}");
                AddLine(fileStream, level, string.Empty);

                AddLine(fileStream, level, "type DataIndex interface {");
                level++;
                AddLine(fileStream, level, "string | int");
                level--;
                AddLine(fileStream, level, "}");
                AddLine(fileStream, level, string.Empty);

                AddLine(fileStream, level, "type DBaseData[T DataIndex] interface {");
                level++;
                AddLine(fileStream, level, "GetID() T");
                level--;
                AddLine(fileStream, level, "}");
            }
        }

        public static void GenerateMapperFile(string packageName, IReadOnlyList<TableInfo> tableInfos, string outputPath)
        {
            string fileName = "mapper.go";
            string fullPath = Path.Combine(outputPath, fileName);
            // 如果文件存在，先删除
            if (File.Exists(fullPath))
                File.Delete(fullPath);
            using (FileStream fileStream = File.Create(fullPath))
            {
                int level = 0;
                AddLine(fileStream, level, $"package {packageName}");
                AddLine(fileStream, level, string.Empty);

                AddLine(fileStream, level, "import \"reflect\"");
                AddLine(fileStream, level, string.Empty);

                AddLine(fileStream, level, "var itemToList = make(map[reflect.Type]reflect.Type)");
                AddLine(fileStream, level, string.Empty);

                AddLine(fileStream, level, "func InitMapper() {");
                level++;
                foreach (var tableInfo in tableInfos)
                    AddLine(fileStream, level, $"registerItemToList(D{tableInfo.TableName}{{}}, D{tableInfo.TableName}List{{}})");
                level--;
                AddLine(fileStream, level, "}");
                AddLine(fileStream, level, string.Empty);

                AddLine(fileStream, level, "func registerItemToList(item, list interface{}) {");
                level++;
                AddLine(fileStream, level, "itemToList[reflect.TypeOf(item)] = reflect.TypeOf(list)");
                level--;
                AddLine(fileStream, level, "}");
                AddLine(fileStream, level, string.Empty);

                AddLine(fileStream, level, "func GetListType(item reflect.Type) reflect.Type {");
                level++;
                AddLine(fileStream, level, "return itemToList[item]");
                level--;
                AddLine(fileStream, level, "}");
            }
        }
    }
}
