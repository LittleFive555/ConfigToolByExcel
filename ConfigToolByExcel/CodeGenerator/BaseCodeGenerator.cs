using System.Text;

namespace ConfigToolByExcel.CodeGenerator
{
    internal class BaseCodeGenerator
    {
        protected const int SpaceCountPerLevel = 4;

        protected static void AddLine(FileStream fs, int level, string value)
        {
            StringBuilder lineStr = new StringBuilder();
            for (int i = 0; i < level * SpaceCountPerLevel; i++)
                lineStr.Append(" ");
            lineStr.Append(value);
            lineStr.Append("\r\n");

            byte[] info = new UTF8Encoding(true).GetBytes(lineStr.ToString());
            fs.Write(info, 0, info.Length);
        }
    }
}
