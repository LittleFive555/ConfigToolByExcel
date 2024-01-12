// See https://aka.ms/new-console-template for more information
using ReadExcel;

// 步骤1.生成代码文件
var classes = ClassReader.CollectClassesInfo("C:\\Users\\LC-MZHANGXI\\Desktop\\111.xlsx");
if (classes != null)
{
    foreach (var classInfo in classes)
        ClassFileGenerator.GenerateClassFile(classInfo, "D:\\UnityProject\\ExcelData\\Assets\\Scripts\\Numeric");
}

// 步骤2.将配置数据转换为json文件
var datas = ClassReader.CollectNumeric("C:\\Users\\LC-MZHANGXI\\Desktop\\111.xlsx");
if (datas != null)
{
    foreach (var data in datas)
    {
        NumericFileGenerator.GenerateNumericFile(data.Key, data.Value, "D:\\UnityProject\\ExcelData\\Assets\\Numeric");
    }
}