using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class StudentNumber:Styles
    {
        public StudentNumber(Word.Document doc, Word.Application app):base(doc,app)
        {

        }

        /*
         * Checks that the document contains the medical/dentistry number in header OR footer
         */
        public bool studentNumberTest()
        {
            bool numberOK = false;
            foreach (Word.Section s in doc.Sections)
            {
                Word.HeadersFooters headers = s.Headers;
                Word.HeaderFooter head = headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                bool ishead = head.IsHeader;
                String text = head.Range.Text;
                bool isMatch = System.Text.RegularExpressions.Regex.IsMatch(text, "(?:M|D)\\d{4}");
                if (!isMatch)
                {
                    numberOK = false;
                }
                else
                {
                    return true;
                }
            }
            if (!numberOK)//if not found in the header, search in footer
            {
                foreach (Word.Section s in doc.Sections)
                {
                    Word.HeadersFooters headers = s.Footers;
                    Word.HeaderFooter head = headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    bool ishead = head.IsHeader;
                    String text = head.Range.Text;
                    bool isMatch = System.Text.RegularExpressions.Regex.IsMatch(text, "(?:M|D)\\d{4}");
                    if (!isMatch)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        /*
         * Checks for the nine digit student number
         */
        public bool studentNumberOrName()
        {
            bool headerOK = false;
            foreach (Word.Section s in doc.Sections)
            {
                Word.HeadersFooters headers = s.Headers;
                foreach (Word.HeaderFooter header in headers)
                {
                    String text = header.Range.Text;
                    bool isMatchNumber = System.Text.RegularExpressions.Regex.IsMatch(text, "\\d{8,9}");
                    bool isMatchName = System.Text.RegularExpressions.Regex.IsMatch(text, "[a-zA-Z]\\d{7}");

                    if (isMatchNumber || isMatchName)
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
