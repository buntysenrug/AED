using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class CaptionRefTest:Styles
    {
        public CaptionRefTest(Word.Document doc, Word.Application app)
            : base(doc, app)
        {

        }
    }
}
