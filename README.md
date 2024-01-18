# README

ConfigToolByExcel是一个用于Unity的Excel配置表工具，可以将按照一定格式配置的Excel表格进行解析，生成C#类代码，并将数据转换为Json格式的文件，以供Unity项目读取。

## 依赖

项目直接依赖的两个包为：

- DocumentFormat.OpenXML 3.0.1

- System.Text.Json 8.0.1

## 项目结构

项目根目录下有两个文件夹`ConfigToolByExcel`和`ConfigToolByExcelUnityEditor`，分别为两个C#项目：

- `ConfigToolByExcel`文件夹下为配置表工具的实现
- `ConfigToolByExcelUnityEditor`文件夹下为Unity中的编辑器实现，用于调用配置表工具。

## 使用方式

### 准备数据

本项目使用的Excel表格配置格式如下：

| 导出标志 | *      | *           |               | *           |
| -------- | ------ | ----------- | ------------- | ----------- |
| 字段名   | NID    | FieldName1  | FieldName2    | FieldName3  |
| 数据类型 | int    | float       | string        | int[]       |
| 备注     | 唯一ID | 自定义字段1 | 自定义字段2   | 自定义字段3 |
| 默认值   | 1      | 2.5         | 无            |             |
| *        | 1      | 1.2         | 我是一个...   | 10#120      |
|          | 2      | 22.3        | My Name Is... | 111#222#333 |
| *        | 3      | 10          | 不懂          | 125         |

一个Excel文件（或称作工作簿Workbook）包含多个工作表（Worksheet），工作表即Excel程序下方的页签（Sheet1、Sheet2...），本项目使用工作表名作为**数据类的类名**，保持大小写。一个工作表会导出一个实体类，一份Json数据文件；当一个Excel文件包含多个工作表时，则会导出工作表数量个实体类，工作表数量个Json数据文件。

- 第一行为字段的**导出标志**，以`*`为标志，若填充`*`，则在导出代码文件时会生成该字段，导出Json文件时会输出该字段的数据；如果为空或为其他值，则导出的代码文件中不会包含该字段，导出的Json文件中也不会含有该字段的值。

- 第二行为**字段名**，最终生成的类代码中的字段名和这里的书写一致，**保持大小写**。

- 第三行为**数据类型**，表示其所在列的对应字段的数据类型，目前支持short、int、long、float、double、string及其对应的一维数组short[]、int[]、long[]、float[]、double[]、string[]。

- 第四行为**备注**，用于描述其所在列的字段，或做一些标记，仅供配表时阅读，不会导出。

- 第五行为**默认值**，当下方某个字段值为空时，会使用其所在列的默认值做填充。

- 从第六行开始，就是**具体的数据配置**，每一行代表一条数据，在读取后时则体现为一个对象。每一行的第一个单元格同样为导出标志，若填充`*`，则在导出Json文件时会将该行数据做输出；如果为空或为其他值，则导出的Json文件不会包含该行数据。

  在配置一维数组时，使用#作数组元素的分隔符。

表中第二列的**NID字段列为固定列**，每个工作表中都**必须含有该列并使用导出符号**，并且该列中配置的数字**不可重复**，作为检索数据的唯一ID。

### 导入工具

#### 直接使用源代码（需联网操作）

1. 在Unity项目中导入DocumentFormat.OpenXML和System.Text.Json的dll及其依赖的dll。

   这里推荐先在Unity中安装[GlitchEnzo/*NuGetForUnity*](https://github.com/GlitchEnzo/NuGetForUnity)。下载NuGetForUnity的最新.unitypackage文件，导入Unity后，搜索并安装OpenXML和System.Text.Json。

2. 下载本项目源代码，将ConfigToolByExcel文件夹下`.cs`代码全部拖入Unity项目的任意位置。

3. 如果想使用自带的可视化工具使用界面，则在Unity项目下新建Editor目录，将ConfigToolByExcelUnityEditor下的所有`.cs`代码全部拖入Editor目录下（UnityDll文件夹不需要拖入）

#### 使用导出的dll

1. 下载最新的Release包，将解压出的ConfigToolByExcel文件夹直接拖入Unity项目的任意位置。
2. 建议将导入的所有.dll通过Inspector面板修改为只在Editor下包含，即Select platform for plugin下，取消Any Platform，然后只保留Editor的选项。

### 生成代码及数据

#### API调用

1. 调用`Generator.GenerateClass()`方法生成类代码。

   ```c#
   /// <summary>
   /// 生成代码文件
   /// </summary>
   /// <param name="excelFilePath">Excel文件所在目录的路径</param>
   /// <param name="codeOutputFolderPath">生成的代码文件的输出路径</param>
   /// <param name="namespaceString">生成代码的命名空间，如果为空，则没有命名空间</param>
   public static void GenerateClass(string excelFilePath, string codeOutputFolderPath, string namespaceString)
   ```
   - excelFilePath为Excel，Excel文件所在的文件夹绝对路径，调用时会将该路径下所有的Excel文件（.xlsx后缀）进行遍历和处理。
   - codeOutputFolderPath，代码文件的输出绝对路径，建议为Unity项目Assets文件夹下自定义路径，**需要事先创建好文件夹**。
   - namespaceString，生成代码的实体类所属的命名空间，如果为空，则没有命名空间。

2. 调用`Generator.GenerateData()`方法生成Json文件。

   ```c#
   /// <summary>
   /// 将配置数据转换为json文件
   /// </summary>
   /// <param name="excelFilePath">Excel文件所在目录的路径</param>
   /// <param name="dataOutputFolderPath">生成的数据Json文件的输出路径</param>
   public static void GenerateData(string excelFilePath, string dataOutputFolderPath)
   ```
   - excelFilePath为Excel，Excel文件所在的文件夹绝对路径，调用时会将该路径下所有的Excel文件（.xlsx后缀）进行遍历和处理。
   - codeOutputFolderPath，Json数据文件的输出绝对路径，建议为Unity项目Assets/StreamingAssets文件夹下新建一个文件夹，**需要事先创建好文件夹**。
   
   输出的Json数据文件后缀为.num。

#### 使用预制的编辑器窗口

1. 首次使用，需要在Unity项目Assets文件夹下新建一个目录Data，用于存放编辑器窗口的配置信息，
2. 点击Unity菜单栏的Config > Open Config Window菜单项，即可打开编辑器窗口。
3. 按照编辑器字段，依次填写路径等信息，填写后点击保存配置信息，可以单独输出类代码或者Json数据文件，也可以一键输出类代码和数据文件。
4. 在首次使用填写好路径并保存配置后，可以直接点击Unity菜单栏的Config > Generate All菜单项，直接生成类代码和数据文件。

### 读取数据

以下是一段简单的数据读取代码，之后会实现一个较为完备的读取类：

```c#
using ConfigData;
using System;
using System.IO;
using System.Linq;
using UnityEngine;

public class ConfigDataManager<T> where T : BaseData
{
    private const string ConfigClassNamespace = "ReadExcel";

    private const string FolderName = "ConfigData";

    public static T GetData(int NID)
    {
        string typeName = typeof(T).Name;
        string rawText = ReadRawText($"{typeName}.num");
        var type = Type.GetType(string.Format("{0}.N{1}List", ConfigClassNamespace, typeName));
        var obj = JsonUtility.FromJson(rawText, type);
        var property = obj.GetType().GetField("Content");
        T[] content = (T[])property.GetValue(obj);
        return content.Where((data) => data.NID == NID).First();
    }

    private static string ReadRawText(string fileName)
    {
        string readData;
        string fileFullPath = Path.Combine(Application.streamingAssetsPath, FolderName, fileName);
        using (StreamReader sr = File.OpenText(fileFullPath))
        {
            readData = sr.ReadToEnd();
            sr.Close();
        }
        return readData;
    }
}
```

