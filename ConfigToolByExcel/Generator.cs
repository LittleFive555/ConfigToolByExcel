using System.IO;

namespace ConfigToolByExcel
{
    public class Generator
    {
        /// <summary>
        /// 生成代码文件
        /// </summary>
        /// <param name="excelFilePath">Excel文件所在目录的路径</param>
        /// <param name="codeOutputFolderPath">生成的代码文件的输出路径</param>
        /// <param name="namespaceString">生成代码的命名空间，如果为空，则没有命名空间</param>
        public static void GenerateClass(string excelFilePath, string codeOutputFolderPath, string namespaceString)
        {
            var fileFullPaths = Directory.GetFiles(excelFilePath);
            foreach (var fullPath in fileFullPaths)
            {
                if (!fullPath.EndsWith(".xlsx"))
                    continue;

                var classes = ClassReader.CollectClassesInfo(fullPath);
                if (classes != null)
                {
                    foreach (var classInfo in classes)
                        ClassFileGenerator.GenerateClassFile(namespaceString, classInfo, codeOutputFolderPath);
                }
            }
        }

        /// <summary>
        /// 将配置数据转换为json文件
        /// </summary>
        /// <param name="excelFilePath">Excel文件所在目录的路径</param>
        /// <param name="dataOutputFolderPath">生成的数据Json文件的输出路径</param>
        public static void GenerateData(string excelFilePath, string dataOutputFolderPath)
        {
            var fileFullPaths = Directory.GetFiles(excelFilePath);
            foreach (var fullPath in fileFullPaths)
            {
                if (!fullPath.EndsWith(".xlsx"))
                    continue;

                var datas = ClassReader.CollectData(fullPath);
                if (datas != null)
                {
                    foreach (var data in datas)
                        DataFileGenerator.GenerateDataFile(data.Key, data.Value, dataOutputFolderPath);
                }
            }
        }
    }
}