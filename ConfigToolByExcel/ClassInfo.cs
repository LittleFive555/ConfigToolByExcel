namespace ConfigToolByExcel
{
    internal struct ClassInfo
    {
        public string ClassName;
        public string ParentClassName;
        public bool IsGeneric;
        public IReadOnlyList<string> GenericTypeList;
        public IReadOnlyList<FieldInfo> Fields;
    }

    internal struct FieldInfo
    {
        public string Name;
        public string Type;
    }
}
