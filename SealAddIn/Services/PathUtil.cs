using System.IO;
using System.Linq;

namespace SealAddIn.Services
{
    internal static class PathUtil
    {
        /// <summary>ファイル名として使用できない文字を "_" に置換します。</summary>
        public static string SanitizeFileNameSegment(string value)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var chars = value.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
            return new string(chars);
        }
    }
}
