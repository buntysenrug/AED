using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class FootNoteStyle:Styles
    {
        public FootNoteStyle(Word.Document doc, Word.Application app)
            : base(doc,app)
        {

        }

        public bool footnoteStyleUsedTest(List<String> footnotequotes)
        {
            if (footnotequotes.Count > 0)
            {
                return false;
            }
            return true;
        }

    }
}
