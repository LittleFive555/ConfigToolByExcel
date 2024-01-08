// See https://aka.ms/new-console-template for more information
using ReadExcel;

var classes = ClassReader.CollectClassesInfo("C:\\Users\\LC-MZHANGXI\\Desktop\\111.xlsx");
foreach (var classInfo in classes)
    ClassFileGenerator.GenerateClassFile(classInfo);
