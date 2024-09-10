using Markdig;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Utilities
{
    public static class StringTool
    {
        public static List<string> GetLines(string path, bool useDecoding = false)
        {
            if (File.Exists(path))
            {
                var lines = File.ReadAllLines(path);

                if (lines.Length == 0)
                {
                    return new List<string>();
                }

                if (useDecoding)
                {
                    return lines.AsParallel().Select(StringCipher.Decode).ToList();
                }
                else
                {
                    return lines.ToList();
                }
            }

            return new List<string>();
        }

        public static bool IsNull(string text)
        {
            if (string.IsNullOrWhiteSpace(text) || string.IsNullOrEmpty(text))
            {
                return true;
            }

            return false;
        }

        public static string AsPlainText(string markdown)
        {
            return Markdown.ToPlainText(markdown).Trim();
        }
    }
}
