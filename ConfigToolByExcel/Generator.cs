namespace ConfigToolByExcel
{
    public class Generator
    {
        // 生成代码文件
        public static void GenerateClass(string excelFilePath, string codeOutputFolderPath)
        {
            var classes = ClassReader.CollectClassesInfo(excelFilePath);
            if (classes != null)
            {
                foreach (var classInfo in classes)
                    ClassFileGenerator.GenerateClassFile(classInfo, codeOutputFolderPath);
            }
        }

        // 将配置数据转换为json文件
        public static void GenerateData(string excelFilePath, string dataOutputFolderPath)
        {
            var datas = ClassReader.CollectNumeric(excelFilePath);
            if (datas != null)
            {
                foreach (var data in datas)
                    NumericFileGenerator.GenerateNumericFile(data.Key, data.Value, dataOutputFolderPath);
            }
        }
    }
}