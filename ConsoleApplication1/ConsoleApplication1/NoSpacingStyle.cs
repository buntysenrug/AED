using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class NoSpacingStyle:Styles
    {
        public NoSpacingStyle(Word.Document doc)
            : base(doc)
        {

        }

        public bool noSpacingStyleUsedTest()
        {
            foreach (Word.Style s in doc.Styles)
            {
                if (s.NameLocal.Equals("No Spacing"))
                {
                    return !s.InUse;
                }
            }
            return true;
        }
    }
}
