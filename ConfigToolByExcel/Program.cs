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

            // C# 代码生成命令
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
            Command csharpCommand = new Command("csharp", "Generate C# classes.")
            {
                excelDirectoryOption,
                codeOutDirectoryOption,
                codeNamespaceOption,
            };
            csharpCommand.SetAction(parseResult =>
            {
                DirectoryInfo excel = parseResult.GetRequiredValue(excelDirectoryOption);
                DirectoryInfo codeOut = parseResult.GetRequiredValue(codeOutDirectoryOption);
                string codeNamespace = parseResult.GetRequiredValue(codeNamespaceOption);
                Generator.GenerateClass(excel.FullName, codeOut.FullName, codeNamespace);
            });

            // JSON 文件生成命令
            Option<DirectoryInfo> jsonOutDirectoryOption = new Option<DirectoryInfo>("--json-out")
            {
                Description = "The directory of generated JSON files.",
                Required = true,
            };
            Command jsonCommand = new Command("json", "Generate JSON files.")
            {
                excelDirectoryOption,
                jsonOutDirectoryOption,
            };
            jsonCommand.SetAction(parseResult =>
            {
                DirectoryInfo excel = parseResult.GetRequiredValue(excelDirectoryOption);
                DirectoryInfo jsonOut = parseResult.GetRequiredValue(jsonOutDirectoryOption);
                Generator.GenerateData(excel.FullName, jsonOut.FullName);
            });

            RootCommand rootCommand = new RootCommand("Excel to json and code generator");
            rootCommand.Subcommands.Add(csharpCommand);
            rootCommand.Subcommands.Add(jsonCommand);

            ParseResult parseResult = rootCommand.Parse(args);
            return parseResult.Invoke(); // 如果有错误会返回错误码，并且显示帮助信息
        }
    }
}
