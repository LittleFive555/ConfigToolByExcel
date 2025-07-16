using System.IO;
using System.Text;

namespace ConfigToolByExcel
{
    internal class CodeFileGenerator
    {
        private const int SpaceCountPerLevel = 4;

        public static void GenerateCSharpFile(string namespaceString, ClassInfo classInfo, string outputPath, bool generaceList)
        {
            string fileName = string.Format("{0}.cs", classInfo.ClassName);
            string fullPath = Path.Combine(outputPath, fileName);
            // 如果文件存在，先删除
            if (File.Exists(fullPath))
                File.Delete(fullPath);

            using (FileStream fileStream = File.Create(fullPath))
            {
                int level = 0;
                AddLine(fileStream, level, "using System;");
                AddLine(fileStream, level, string.Empty);

                // namespace start
                if (!string.IsNullOrEmpty(namespaceString))
                {
                    AddLine(fileStream, level, string.Format("namespace {0}", namespaceString));
                    AddLine(fileStream, level, "{");
                    level++;
                }
                GenerateCSharpClass(classInfo, fileStream, ref level);
                AddLine(fileStream, level, string.Empty);

                if (generaceList)
                {
                    // class start
                    AddLine(fileStream, level, string.Format("public class {0}List", classInfo.ClassName));
                    AddLine(fileStream, level, "{");

                    level++;
                    // property start
                    AddLine(fileStream, level, string.Format("public {0}[] Content;", classInfo.ClassName));
                    // property end

                    level--;
                    AddLine(fileStream, level, "}");
                    // class end
                }

                level--;
                // namespace end
                if (!string.IsNullOrEmpty(namespaceString))
                    AddLine(fileStream, level, "}");
            }
        }

        public static void GenerateCSharpClass(ClassInfo classInfo, FileStream fileStream, ref int level)
        {
            // class start
            AddLine(fileStream, level, "[Serializable]");

            StringBuilder classFullNameBuilder = new StringBuilder();
            classFullNameBuilder.Append(classInfo.ClassName);
            if (classInfo.IsGeneric && classInfo.GenericTypeList != null && classInfo.GenericTypeList.Count > 0)
            {
                classFullNameBuilder.Append("<");
                classFullNameBuilder.Append(string.Join(", ", classInfo.GenericTypeList));
                classFullNameBuilder.Append(">");
            }
            if (string.IsNullOrEmpty(classInfo.ParentClassName))
                AddLine(fileStream, level, string.Format("public class {0}", classFullNameBuilder.ToString()));
            else
                AddLine(fileStream, level, string.Format("public class {0} : {1}", classFullNameBuilder.ToString(), classInfo.ParentClassName));
            AddLine(fileStream, level, "{");
            level++;

            // property start
            if (classInfo.Fields != null)
            {
                foreach (FieldInfo fieldInfo in classInfo.Fields)
                    AddLine(fileStream, level, $"public {fieldInfo.Type} {fieldInfo.Name};");
            }
            // property end

            level--;
            AddLine(fileStream, level, "}");
        }

        public static void GenerateGoFile(string packageName, ClassInfo classInfo, string outputPath)
        {
            string fileName = string.Format("{0}.go", classInfo.ClassName);
            string fullPath = Path.Combine(outputPath, fileName);
            // 如果文件存在，先删除
            if (File.Exists(fullPath))
                File.Delete(fullPath);
            using (FileStream fileStream = File.Create(fullPath))
            {
                int level = 0;
                AddLine(fileStream, level, string.Format("package {0}", packageName));
                AddLine(fileStream, level, string.Empty);
                // struct start
                AddLine(fileStream, level, string.Format("type {0} struct {{", classInfo.ClassName));
                level++;
                // property start
                if (classInfo.Fields != null)
                {
                    foreach (FieldInfo fieldInfo in classInfo.Fields)
                        AddLine(fileStream, level, $"{fieldInfo.Name} {fieldInfo.Type}");
                }
                // property end
                level--;
                AddLine(fileStream, level, "}");

                AddLine(fileStream, level, string.Empty);

                AddLine(fileStream, level, string.Format("type {0}List struct {{", classInfo.ClassName));
                level++;
                AddLine(fileStream, level, string.Format("Content []{0}", classInfo.ClassName));
                level--;
                AddLine(fileStream, level, "}");

            }
        }

        private static void AddLine(FileStream fs, int level, string value)
        {
            StringBuilder lineStr = new StringBuilder();
            for (int i = 0; i < level * SpaceCountPerLevel; i++)
                lineStr.Append(" ");
            lineStr.Append(value);
            lineStr.Append("\r\n");

            byte[] info = new UTF8Encoding(true).GetBytes(lineStr.ToString());
            fs.Write(info, 0, info.Length);
        }
    }
}
