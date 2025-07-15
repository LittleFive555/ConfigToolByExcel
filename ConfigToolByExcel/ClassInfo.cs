namespace ConfigToolByExcel
{
    internal struct ClassInfo
    {
        public string ClassName;
        public string ParentClassName;
        public IReadOnlyList<PropertyInfo> Properties;
    }

    internal struct PropertyInfo
    {
        public string Name;
        public string Type;
    }
}
