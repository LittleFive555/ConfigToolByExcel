namespace ConfigToolByExcel.CodeGenerator
{
    internal struct TableInfo
    {
        public string TableName;
        public string IDType;
        public IReadOnlyList<FieldInfo> Fields;
    }

    internal struct FieldInfo
    {
        public string Name;
        public string Type;
    }
}
