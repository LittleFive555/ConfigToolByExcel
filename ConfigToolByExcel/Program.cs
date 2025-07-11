namespace ConfigToolByExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
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
            Generator.GenerateClass(excelPath, codeOutputPath, codeNamaspace);
        }

        public static void DoGenerateData(string excelPath, bool isPathRelative, string dataOutputPath)
        {
            if (isPathRelative)
                dataOutputPath = Path.Combine(Environment.CurrentDirectory, dataOutputPath);
            Generator.GenerateData(excelPath, dataOutputPath);
        }
    }
}
