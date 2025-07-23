using ConfigToolByExcel.CodeGenerator;

namespace ConfigToolByExcel
{
    public class Commands
    {
        /// <summary>
        /// 生成代码文件
        /// </summary>
        /// <param name="excelFilePath">Excel文件所在目录的路径</param>
        /// <param name="codeOutputFolderPath">生成的代码文件的输出路径</param>
        /// <param name="namespaceString">生成代码的命名空间，如果为空，则没有命名空间</param>
        public static void GenerateCSharpFiles(string excelFilePath, string codeOutputFolderPath, string namespaceString)
        {
            Directory.CreateDirectory(codeOutputFolderPath);

            CSharpCodeGenerator.GenerateBaseClassFile(namespaceString, codeOutputFolderPath);

            var fileFullPaths = Directory.GetFiles(excelFilePath);
            foreach (var fullPath in fileFullPaths)
            {
                if (!fullPath.EndsWith(".xlsx"))
                    continue;

                var tables = ExcelReader.CollectTableInfo(fullPath);
                if (tables != null)
                {
                    foreach (var table in tables)
                        CSharpCodeGenerator.GenerateCSharpFile(namespaceString, table, codeOutputFolderPath);
                }
            }
        }

        public static void GenerateGoFiles(string excelFilePath, string codeOutputFolderPath, string packageName)
        {
            Directory.CreateDirectory(codeOutputFolderPath);

            GoCodeGenerator.GenerateInterfaceFile(packageName, codeOutputFolderPath);

            var fileFullPaths = Directory.GetFiles(excelFilePath);
            List<TableInfo> classes = new List<TableInfo>();
            foreach (var fullPath in fileFullPaths)
            {
                if (!fullPath.EndsWith(".xlsx"))
                    continue;

                var someTables = ExcelReader.CollectTableInfo(fullPath);
                if (someTables == null)
                    continue;

                classes.AddRange(someTables);
                foreach (var classInfo in classes)
                    GoCodeGenerator.GenerateGoFile(packageName, classInfo, codeOutputFolderPath);
            }

            GoCodeGenerator.GenerateMapperFile(packageName, classes, codeOutputFolderPath);
        }

        /// <summary>
        /// 将配置数据转换为json文件
        /// </summary>
        /// <param name="excelFilePath">Excel文件所在目录的路径</param>
        /// <param name="dataOutputFolderPath">生成的数据Json文件的输出路径</param>
        public static void GenerateData(string excelFilePath, string dataOutputFolderPath)
        {
            Directory.CreateDirectory(dataOutputFolderPath);
            var fileFullPaths = Directory.GetFiles(excelFilePath);
            foreach (var fullPath in fileFullPaths)
            {
                if (!fullPath.EndsWith(".xlsx"))
                    continue;

                var datas = ExcelReader.CollectData(fullPath);
                if (datas != null)
                {
                    foreach (var data in datas)
                        DataFileGenerator.GenerateDataFile(data.Key, data.Value, dataOutputFolderPath);
                }
            }
        }
    }
}