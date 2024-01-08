using System.Text.Json.Serialization;

namespace ReadExcel
{
    public class BaseData
    {
        [JsonPropertyOrder(-1)]
        public int NID { get; set; }
    }
}
