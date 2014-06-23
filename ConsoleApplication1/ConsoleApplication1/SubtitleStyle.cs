using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class SubtitleStyle:Styles
    {
        public SubtitleStyle(Word.Document doc)
            : base(doc)
        {

        }

        public bool subTitileStyleUsedTest()
        {
            return !style_name.Contains("Subtitle");
        }
    }
}
