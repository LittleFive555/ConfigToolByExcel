using UnityEngine;

public class ConfigWindowData : ScriptableObject
{
    [Header("配置表路径")]
    [SerializeField]
    public string ExcelPath;

    [Header("代码所在命名空间")]
    [SerializeField]
    public string CodeNamespace;

    [Header("代码输出路径是否为相对路径")]
    [SerializeField]
    public bool IsCodeOutputPathRelative;

    [Header("代码输出路径")]
    [SerializeField]
    public string CodeOutputPath;

    [Header("数据输出路径是否为相对路径")]
    [SerializeField]
    public bool IsDataOutputPathRelative;

    [Header("数据输出路径")]
    [SerializeField]
    public string DataOutputPath;
}
