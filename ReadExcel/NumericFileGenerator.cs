using System.Text;
using System.Text.Json;

namespace ReadExcel
{
    internal class NumericFileGenerator
    {
        private static readonly string OutputPath = "C:\\Users\\LC-MZHANGXI\\Desktop\\NumericData";

        public static void GenerateNumericFile(string fileName, IReadOnlyList<BaseData> datas)
        {
            string fullPath = Path.Combine(OutputPath, string.Format("{0}.num", fileName));
            if (File.Exists(fullPath))
                File.Delete(fullPath);
            StringBuilder sb = new StringBuilder();
            foreach (var data in datas)
                sb.AppendLine(JsonSerializer.Serialize(data, data.GetType()));
            File.WriteAllText(fullPath, sb.ToString());
        }
    }
}
