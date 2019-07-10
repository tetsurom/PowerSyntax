using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerSyntax
{
    [DebuggerDisplay("({Begin}..{Begin + Length}) {Text} : {Color}")]
    public class StyleRange
    {
        public string Text { get; set; }
        public StyleRange() { }
        public int Begin { get; set; }
        public int Length { get; set; }
        public Color? Color { get; set; }
        public Color? BackGroundColor { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public IEnumerable<StyleRange> Children { get; set; }
    }

    public interface IResourceProvider
    {
        IEnumerable<string> GetFiles(string path, string type);
        string ReadAllText(string path);
    }

    public interface ISyntaxHighlighter
    {
        StyleRange Parse(string code, string language, string theme);
    }
}
