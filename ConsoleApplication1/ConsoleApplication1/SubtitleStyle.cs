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
            foreach (Word.Style s in doc.Styles)
            {
                if (s.NameLocal.Equals("Subtitle"))
                {
                    return !s.InUse;
                }
            }
            return true;
        }
    }
}
