using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class ListParagraph:Styles
    {
         public ListParagraph(Word.Document doc,Word.Application app)
            : base(doc,app)
        {

        }

        public bool listParaNumUsed()
        {
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                Word.Style s = p.get_Style();
                if (s.NameLocal.Equals("List Paragraph"))
                {
                    Word.ListFormat list = p.Range.ListFormat;
                    return System.Text.RegularExpressions.Regex.IsMatch(list.ListString, "\\d");
                }
            }
            return false;
        }

        public bool listParaBulletedUsed()
        {
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                Word.Style s = p.get_Style();
                if (s.NameLocal.Equals("List Paragraph"))
                {
                    Word.ListFormat list = p.Range.ListFormat;
                    return System.Text.RegularExpressions.Regex.IsMatch(list.ListString, ".");
                }
            }
            return false;
        }
    }
}
