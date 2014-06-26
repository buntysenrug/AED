using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class HeaderStyle:Styles
    {
        public HeaderStyle(Word.Document doc, Word.Application app)
            : base(doc,app)
        {

        }

        public bool headerStyleUsedTest()
        {
            bool headerOK = false;
            foreach (Word.Section s in doc.Sections)
            {
                Word.HeadersFooters headers = s.Headers;
                foreach (Word.HeaderFooter header in headers)
                {
                    bool match = System.Text.RegularExpressions.Regex.IsMatch(header.Range.Text, "\\w+");
                    if (match)
                    {
                        headerOK = true;
                        break;
                    }
                }
                if (headerOK)
                {
                    break;
                }
            }
            return headerOK;
        }


    }
}
