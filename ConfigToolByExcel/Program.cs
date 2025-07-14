using System.CommandLine;

namespace ConfigToolByExcel
{
    internal class Program
    {
        static int Main(string[] args)
        {
            Option<DirectoryInfo> excelDirectoryOption = new Option<DirectoryInfo>("--excel")
            {
                Description = "The directory containing Excel files.",
                Required = true,
            };

            Option<DirectoryInfo> codeOutDirectoryOption = new Option<DirectoryInfo>("--code-out")
            {
                Description = "The directory of generated code files.",
                Required = true,
            };

            Option<string> codeNamespaceOption = new Option<string>("--code-namespace")
            {
                Description = "The namespace for the generated code.",
                Required = true,
            };

            Option<DirectoryInfo> jsonOutDirectoryOption = new Option<DirectoryInfo>("--json-out")
            {
                Description = "The directory of generated JSON files.",
                Required = true,
            };

            RootCommand rootCommand = new RootCommand("Excel to json and class generator")
            {
                excelDirectoryOption,
                codeOutDirectoryOption,
                codeNamespaceOption,
                jsonOutDirectoryOption,
            };

            rootCommand.SetAction(parseResult =>
            {
                DirectoryInfo excel = parseResult.GetRequiredValue(excelDirectoryOption);
                DirectoryInfo codeOut = parseResult.GetRequiredValue(codeOutDirectoryOption);
                string codeNamespace = parseResult.GetRequiredValue(codeNamespaceOption);
                DirectoryInfo jsonOut = parseResult.GetRequiredValue(jsonOutDirectoryOption);
                Generator.GenerateClass(excel.FullName, codeOut.FullName, codeNamespace);
                Generator.GenerateData(excel.FullName, jsonOut.FullName);
                return 0;
            });
            ParseResult parseResult = rootCommand.Parse(args);
            return parseResult.Invoke(); // 如果有错误会返回错误码，并且显示帮助信息
        }
    }
}
