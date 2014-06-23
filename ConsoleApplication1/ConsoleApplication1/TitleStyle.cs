using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class TitleStyle:Styles
    {
        public TitleStyle(Word.Document doc)
            : base(doc)
        {

        }

        public bool runTitleUsed()
        {
            return true;
        }
    }
}
