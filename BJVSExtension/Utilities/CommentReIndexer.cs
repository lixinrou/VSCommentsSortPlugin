using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BJVSExtension.Utilities
{
    class CommentReIndexer
    {
        public const string DoubleSlash = "//";
        public int TabSize { get; }
        public string LineEnding { get; }

        private readonly List<string> lines;
        private readonly List<string> linesWithoutTabs;

        public CommentReIndexer(IEnumerable<string> lines, int tabSize, string lineEnding)
        {
            TabSize = tabSize;
            LineEnding = lineEnding;

            this.lines = lines.ToList();
            this.linesWithoutTabs = lines.Select(x => x.Replace("\t", new string(' ', TabSize))).ToList();
        }

        public string GetText()
        {
            string text = string.Join(LineEnding, GetLines()) + LineEnding;
            return text;
        }

        private IEnumerable<string> GetLines()
        {
            
            int commentIndex = 1;

            List<string> newLines = new List<string>();
            for (int i = 0; i < lines.Count(); i++)
            {
                string line = lines[i];

                int index = line.LastIndexOf(DoubleSlash);
                
                if (index <= -1)
                {
                    // Add unchanged line
                    newLines.Add(line);
                }
                else
                {
                    // 注释
                    string comment = line.Substring(index);
                    string output = "";
                    // 如果注释符合 // Num.格式
                    string pattern = @"(?<=// )[1-9]\d*\b";
                    if (Regex.IsMatch(comment, pattern))
                    {
                        Regex reg = new Regex(pattern);
                        output = reg.Replace(comment, commentIndex.ToString());
                        commentIndex += 1;
                    

                        // 注释前的正文
                        string subString = line.Substring(0, index);

                        string newLine = subString + output;
                        
                        newLines.Add(newLine);
                    }
                    else
                    {
                        newLines.Add(line);
                    }
                }
            }
            return newLines;
        }
    }
}
