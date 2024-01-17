using UnityEditor;
using UnityEngine;

public class ConfigWindow : EditorWindow
{
    private static ConfigWindow m_instance = null;
    private static ConfigWindow Instance
    {
        get
        {
            if (m_instance == null)
                m_instance = GetWindow<ConfigWindow>(false, WindowName, true);
            return m_instance;
        }
    }

    public const string DataPath = "Assets/Data/ConfigWindowData.asset";

    private const float AreaSpace = 10;

    private const string WindowName = "Config Window";
    private const string WindowNameDirty = "Config Window *";

    private ConfigWindowData m_configWindowData;

    public static void OpenWindow()
    {
        Instance.Show();
    }

    private void OnGUI()
    {
        GetOrCreateAsset();

        GUILayout.Space(AreaSpace);
        GUILayout.BeginHorizontal();
        GUILayout.Space(AreaSpace);
        GUILayout.BeginVertical();

        EditorGUI.BeginChangeCheck();

        m_configWindowData.ExcelPath = EditorGUILayout.TextField("Excel·��", m_configWindowData.ExcelPath);

        GUILayout.Space(AreaSpace);

        // ���ɴ���
        m_configWindowData.CodeNamespace = EditorGUILayout.TextField("�������������ռ�", m_configWindowData.CodeNamespace);
        m_configWindowData.IsCodeOutputPathRelative = EditorGUILayout.Toggle("ʹ�����·��", m_configWindowData.IsCodeOutputPathRelative);
        GUILayout.BeginHorizontal();
        m_configWindowData.CodeOutputPath = EditorGUILayout.TextField("�����ļ����·��", m_configWindowData.CodeOutputPath);
        if (EditorGUI.EndChangeCheck())
            Instance.titleContent.text = WindowNameDirty;

        if (GUILayout.Button("���ɴ���", GUILayout.Width(100)))
            ConfigMenuEditor.DoGenerateClass(m_configWindowData.ExcelPath, m_configWindowData.IsCodeOutputPathRelative, m_configWindowData.CodeOutputPath, m_configWindowData.CodeNamespace);
        GUILayout.EndHorizontal();

        GUILayout.Space(AreaSpace);

        EditorGUI.BeginChangeCheck();
        // ��������
        m_configWindowData.IsDataOutputPathRelative = EditorGUILayout.Toggle("ʹ�����·��", m_configWindowData.IsDataOutputPathRelative);

        GUILayout.BeginHorizontal();
        m_configWindowData.DataOutputPath = EditorGUILayout.TextField("�����ļ����·��", m_configWindowData.DataOutputPath);

        if (EditorGUI.EndChangeCheck())
            Instance.titleContent.text = WindowNameDirty;

        if (GUILayout.Button("��������", GUILayout.Width(100)))
            ConfigMenuEditor.DoGenerateData(m_configWindowData.ExcelPath, m_configWindowData.IsDataOutputPathRelative, m_configWindowData.DataOutputPath);
        GUILayout.EndHorizontal();
        GUILayout.Space(AreaSpace);

        if (GUILayout.Button("ȫ������"))
        {
            ConfigMenuEditor.DoGenerateAll(m_configWindowData.ExcelPath,
                m_configWindowData.IsCodeOutputPathRelative, m_configWindowData.CodeOutputPath, m_configWindowData.CodeNamespace,
                m_configWindowData.IsDataOutputPathRelative, m_configWindowData.DataOutputPath);
        }

        GUILayout.Space(AreaSpace);

        if (GUILayout.Button("���洰������"))
        {
            AssetDatabase.SaveAssetIfDirty(m_configWindowData);
            AssetDatabase.Refresh();
            Instance.titleContent.text = WindowName;
        }

        GUILayout.EndVertical();
        GUILayout.Space(AreaSpace);
        GUILayout.EndHorizontal();
    }

    private void GetOrCreateAsset()
    {
        if (m_configWindowData == null)
            m_configWindowData = AssetDatabase.LoadAssetAtPath<ConfigWindowData>(DataPath);
        if (m_configWindowData == null)
        {
            m_configWindowData = (ConfigWindowData)ScriptableObject.CreateInstance(typeof(ConfigWindowData));
            AssetDatabase.CreateAsset(m_configWindowData, DataPath);
        }
    }
}
