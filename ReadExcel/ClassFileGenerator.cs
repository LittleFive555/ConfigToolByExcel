using System.Text;

namespace ReadExcel
{
    internal class ClassFileGenerator
    {
        private const int SpaceCountPerLevel = 4;
        private const string NamespaceStr = "ReadExcel";
        private const string BaseClassName = "BaseData";

        public static void GenerateClassFile(ClassInfo classInfo, string outputPath)
        {
            string fileName = string.Format("{0}.cs", classInfo.ClassName);
            string fullPath = Path.Combine(outputPath, fileName);
            // 如果文件存在，先删除
            if (File.Exists(fullPath))
                File.Delete(fullPath);

            using (FileStream fileStream = File.Create(fullPath))
            {
                AddLine(fileStream, 0, string.Format("namespace {0}", NamespaceStr));
                AddLine(fileStream, 0, "{");
                if (classInfo.ClassName.Equals(BaseClassName))
                    AddLine(fileStream, 1, string.Format("public class {0}", classInfo.ClassName));
                else
                    AddLine(fileStream, 1, string.Format("public class {0} : {1}", classInfo.ClassName, BaseClassName));
                AddLine(fileStream, 1, "{");
                foreach (PropertyInfo fieldInfo in classInfo.Properties)
                    AddLine(fileStream, 2, $"public {fieldInfo.Type} {fieldInfo.Name} {{ get; set; }}");
                AddLine(fileStream, 1, "}");
                AddLine(fileStream, 0, "}");
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
