using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class PageNumber:Styles
    {
        public PageNumber(Word.Document doc, Word.Application app)
            : base(doc, app)
        {

        }

        /*
         * A test to determine if a page number has been used at least once in the doc
         */
        public bool pageNumberTest()
        {

            foreach (Word.Section s in doc.Sections)
            {
                foreach (Word.HeaderFooter foot in s.Footers)
                {
                    String text = foot.Range.Text;
                    Word.PageNumbers page = foot.PageNumbers;
                    int count = page.Count;
                    if (count > 0)
                    {
                        return true;
                    }
                }
            }
            return false;
        }


    }
}
