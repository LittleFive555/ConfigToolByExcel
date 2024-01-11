using System.Text;
using System.Text.Json;

namespace ReadExcel
{
    internal class NumericFileGenerator
    {
        public static void GenerateNumericFile(string fileName, IReadOnlyList<BaseData> datas, string outputPath)
        {
            string fullPath = Path.Combine(outputPath, string.Format("{0}.num", fileName));
            if (File.Exists(fullPath))
                File.Delete(fullPath);
            StringBuilder sb = new StringBuilder();
            foreach (var data in datas)
                sb.AppendLine(JsonSerializer.Serialize(data, data.GetType()));
            File.WriteAllText(fullPath, sb.ToString());
        }
    }
}
