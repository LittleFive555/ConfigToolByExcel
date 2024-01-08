namespace ReadExcel
{
    internal struct ClassInfo
    {
        public string ClassName;
        public IReadOnlyList<FieldInfo> Fields;
    }

    internal struct FieldInfo
    {
        public string Name;
        public string Type;
    }
}
