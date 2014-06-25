using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class FooterStyle:Styles
    {
        public FooterStyle(Word.Document doc)
            : base(doc)
        {

        }

        public bool footerStyleUsedTest()
        {
            bool footerOK = false;
            foreach (Word.Section s in doc.Sections)
            {
                Word.HeadersFooters headers = s.Footers;
                foreach (Word.HeaderFooter footer in headers)
                {
                    bool match = System.Text.RegularExpressions.Regex.IsMatch(footer.Range.Text, "\\w+");
                    if (match)
                    {
                        footerOK = true;
                        break;
                    }
                }
                if (footerOK)
                {
                    break;
                }
            }
            return footerOK;
        }



    }
}
