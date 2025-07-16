namespace ConfigToolByExcel.CodeGenerator
{
    /// <summary>
    /// C#代码生成器
    /// 
    /// <para>生成规则如下：</para>
    /// <para>文件名为Excel表名，每张表生成两个类，一个类（类名为 D{tablename}）为带有表中各个字段的数据类，另一个类（类名为 D{tablename}List）的字段为上一个类的数组。</para>
    /// </summary>
    internal class CSharpCodeGenerator : BaseCodeGenerator
    {
        private const string BaseClassName = "DBaseData";

        public static void GenerateCSharpFile(string namespaceString, TableInfo tableInfo, string outputPath)
        {
            string fileName = string.Format("{0}.cs", tableInfo.TableName);
            string fullPath = Path.Combine(outputPath, fileName);
            // 如果文件存在，先删除
            if (File.Exists(fullPath))
                File.Delete(fullPath);

            using (FileStream fileStream = File.Create(fullPath))
            {
                int level = 0;
                AddLine(fileStream, level, $"using System;");
                AddLine(fileStream, level, string.Empty);

                // namespace start
                if (!string.IsNullOrEmpty(namespaceString))
                {
                    AddLine(fileStream, level, $"namespace {namespaceString}");
                    AddLine(fileStream, level, $"{{");
                    level++;
                }
                GenerateCSharpClass(tableInfo, fileStream, ref level);
                AddLine(fileStream, level, string.Empty);

                // class start
                AddLine(fileStream, level, $"public class D{tableInfo.TableName}List");
                AddLine(fileStream, level, $"{{");

                level++;
                // property start
                AddLine(fileStream, level, $"public D{tableInfo.TableName}[] Content;");
                // property end
                level--;

                AddLine(fileStream, level, $"}}");
                // class end

                // namespace end
                if (!string.IsNullOrEmpty(namespaceString))
                {
                    level--;
                    AddLine(fileStream, level, $"}}");
                }
            }
        }

        public static void GenerateBaseClassFile(string namespaceString, string outputPath)
        {
            string fullPath = Path.Combine(outputPath, "BaseData.cs");
            // 如果文件存在，先删除
            if (File.Exists(fullPath))
                File.Delete(fullPath);

            using (FileStream fileStream = File.Create(fullPath))
            {
                int level = 0;
                //AddLine(fileStream, level, $"using System;");
                //AddLine(fileStream, level, string.Empty);

                // namespace start
                if (!string.IsNullOrEmpty(namespaceString))
                {
                    AddLine(fileStream, level, $"namespace {namespaceString}");
                    AddLine(fileStream, level, $"{{");
                    level++;
                }
                AddLine(fileStream, level, string.Empty);

                // class start
                AddLine(fileStream, level, $"public class {BaseClassName}<TIndex>");
                AddLine(fileStream, level, $"{{");

                level++;
                // property start
                AddLine(fileStream, level, $"public TIndex ID;");
                // property end
                level--;

                AddLine(fileStream, level, $"}}");
                // class end

                level--;
                // namespace end
                if (!string.IsNullOrEmpty(namespaceString))
                    AddLine(fileStream, level, $"}}");
            }
        }

        private static void GenerateCSharpClass(TableInfo tableInfo, FileStream fileStream, ref int level)
        {
            // class start
            AddLine(fileStream, level, $"[Serializable]");

            AddLine(fileStream, level, $"public class D{tableInfo.TableName} : DBaseData<{tableInfo.IDType}>");
            AddLine(fileStream, level, $"{{");
            level++;

            // property start
            if (tableInfo.Fields != null)
            {
                foreach (FieldInfo fieldInfo in tableInfo.Fields)
                {
                    if (fieldInfo.Name == "ID") // 基类包含
                        continue;
                    AddLine(fileStream, level, $"public {fieldInfo.Type} {fieldInfo.Name};"); // TODO 类型转换为C#类型
                }
            }
            // property end

            level--;
            AddLine(fileStream, level, $"}}");
        }
    }
}
