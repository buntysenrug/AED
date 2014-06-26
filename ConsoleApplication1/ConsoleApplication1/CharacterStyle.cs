using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class CharacterStyle:Styles
    {
        public CharacterStyle(Word.Document doc, Word.Application app)
            : base(doc,app)
        {

        }


        public bool characterStyleTest(List<String> characterquotes)
        {
            if (characterquotes.Count > 5)
            {
                return false;
            }
            return true;
        }
    }
}
