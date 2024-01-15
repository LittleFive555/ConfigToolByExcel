namespace ConfigToolByExcel
{
    public class Generator
    {
        public static void Generate(string excelFilePath, string codeOutputFolderPath, string dataOutputFolderPath)
        {
            // 步骤1.生成代码文件
            var classes = ClassReader.CollectClassesInfo(excelFilePath);
            if (classes != null)
            {
                foreach (var classInfo in classes)
                    ClassFileGenerator.GenerateClassFile(classInfo, codeOutputFolderPath);
            }

            // 步骤2.将配置数据转换为json文件
            var datas = ClassReader.CollectNumeric(excelFilePath);
            if (datas != null)
            {
                foreach (var data in datas)
                {
                    NumericFileGenerator.GenerateNumericFile(data.Key, data.Value, dataOutputFolderPath);
                }
            }
        }
    }
}