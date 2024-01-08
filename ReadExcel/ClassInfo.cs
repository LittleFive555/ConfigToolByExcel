namespace ReadExcel
{
    internal struct ClassInfo
    {
        public string ClassName;
        public IReadOnlyList<PropertyInfo> Properties;
    }

    internal struct PropertyInfo
    {
        public string Name;
        public string Type;
    }
}
