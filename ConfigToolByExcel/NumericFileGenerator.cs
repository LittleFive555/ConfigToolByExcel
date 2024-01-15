using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ConfigToolByExcel
{
    internal class NumericFileGenerator
    {
        public static void GenerateNumericFile(string fileName, JsonNode datas, string outputPath)
        {
            string fullPath = Path.Combine(outputPath, string.Format("{0}.num", fileName));
            if (File.Exists(fullPath))
                File.Delete(fullPath);

            var options = new JsonSerializerOptions(JsonSerializerDefaults.General);
            options.WriteIndented = true;
            File.WriteAllText(fullPath, datas.ToJsonString(options));
        }
    }
}
