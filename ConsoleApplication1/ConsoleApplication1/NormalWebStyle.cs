using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class NormalWebStyle:Styles
    {
        public NormalWebStyle(Word.Document doc)
            : base(doc)
        {

        }

        public bool normalWebStyleUsedTest(List<String> normalwebquotes)
        {
            if (normalwebquotes.Count > 0)
            {
                return false;
            }
            return true;
        }


    }
}
