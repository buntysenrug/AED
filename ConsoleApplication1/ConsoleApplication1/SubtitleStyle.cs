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

        public bool subTitileStyleUsedTest(List<String> subtitlequotes)
        {
            return subtitlequotes.Count == 0;
        }
    }
}
