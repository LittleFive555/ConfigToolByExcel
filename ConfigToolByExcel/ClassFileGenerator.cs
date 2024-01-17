using System.IO;
using System.Text;

namespace ConfigToolByExcel
{
    internal class ClassFileGenerator
    {
        private const int SpaceCountPerLevel = 4;
        private const string BaseClassName = "BaseData";

        public static void GenerateClassFile(string namespaceString, ClassInfo classInfo, string outputPath)
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

                // class start
                AddLine(fileStream, level, "[Serializable]");
                if (classInfo.ClassName.Equals(BaseClassName))
                    AddLine(fileStream, level, string.Format("public class {0}", classInfo.ClassName));
                else
                    AddLine(fileStream, level, string.Format("public class {0} : {1}", classInfo.ClassName, BaseClassName));
                AddLine(fileStream, level, "{");
                level++;

                // property start
                foreach (PropertyInfo fieldInfo in classInfo.Properties)
                    AddLine(fileStream, level, $"public {fieldInfo.Type} {fieldInfo.Name};");
                // property end

                level--;
                AddLine(fileStream, level, "}");

                AddLine(fileStream, level, string.Empty);

                // class start
                AddLine(fileStream, level, string.Format("public class N{0}List", classInfo.ClassName));
                AddLine(fileStream, level, "{");

                level++;
                // property start
                AddLine(fileStream, level, string.Format("public {0}[] Content;", classInfo.ClassName));
                // property end

                level--;
                AddLine(fileStream, level, "}");
                // class end

                level--;
                // namespace end
                if (!string.IsNullOrEmpty(namespaceString))
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
