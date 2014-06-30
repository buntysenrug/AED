using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class PageOrientation:Styles
    {
        public PageOrientation(Word.Document doc, Word.Application app)
            : base(doc, app)
        {

        }

        /*
         * Tests if the doc has a page with its orientation set to landscape. 
         */
        public bool landscapePageTest()
        {
            foreach (Word.Section s in doc.Sections)
            {
                if (s.PageSetup.Orientation == Word.WdOrientation.wdOrientLandscape)
                {
                    return true;
                }
            }
            
            return false;

        }
    }
}
