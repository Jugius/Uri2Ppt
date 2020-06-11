using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Uri2Ppt
{
    public sealed class DownloadedItem
    {
        public string Text { get; set; }
        public Uri Hyperlink { get; set; }
        public List<string> Bitmaps { get; set; } = new List<string>();
    }
}
