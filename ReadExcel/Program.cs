// See https://aka.ms/new-console-template for more information
using ReadExcel;
using System.Text.Json;

// 以下两个步骤需要分两次运行，因为步骤2需要依赖于步骤1生成的代码来读取数据并生成json文件。
// 步骤1生成代码后，需要经过编译，再执行步骤2

// TODO 改用类代码和json文件无相互依赖的实现方式，例如直接从excel文件转为json文件

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
        NumericFileGenerator.GenerateNumericFile(data.Key.Name, data.Value, "D:\\UnityProject\\ExcelData\\Assets\\Numeric");
    }
}