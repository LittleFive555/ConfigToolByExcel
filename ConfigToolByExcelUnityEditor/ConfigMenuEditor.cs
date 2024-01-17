using System;
using System.IO;
using UnityEditor;

public class ConfigMenuEditor
{
    [MenuItem("Config/Open Config Window")]
    public static void OpenConfigWindow()
    {
        ConfigWindow.OpenWindow();
    }

    [MenuItem("Config/Generate All")]
    public static void GenerateAll()
    {
        var configEditorData = AssetDatabase.LoadAssetAtPath<ConfigWindowData>(ConfigWindow.DataPath);
        DoGenerateAll(configEditorData.ExcelPath,
            configEditorData.IsCodeOutputPathRelative, configEditorData.CodeOutputPath, configEditorData.CodeNamespace,
            configEditorData.IsDataOutputPathRelative, configEditorData.DataOutputPath);
        AssetDatabase.Refresh();
    }

    public static void DoGenerateAll(string excelPath,
        bool isCodePathRelative, string codeOutputPath, string codeNamaspace,
        bool isDataPathRelative, string dataOutputPath)
    {
        DoGenerateClass(excelPath, isCodePathRelative, codeOutputPath, codeNamaspace);
        DoGenerateData(excelPath, isDataPathRelative, dataOutputPath);
    }

    public static void DoGenerateClass(string excelPath, bool isPathRelative, string codeOutputPath, string codeNamaspace)
    {
        if (isPathRelative)
            codeOutputPath = Path.Combine(Environment.CurrentDirectory, codeOutputPath);
        ConfigToolByExcel.Generator.GenerateClass(excelPath, codeOutputPath, codeNamaspace);
    }

    public static void DoGenerateData(string excelPath, bool isPathRelative, string dataOutputPath)
    {
        if (isPathRelative)
            dataOutputPath = Path.Combine(Environment.CurrentDirectory, dataOutputPath);
        ConfigToolByExcel.Generator.GenerateData(excelPath, dataOutputPath);
    }
}
