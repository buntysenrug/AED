using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class TableOfContentTest:Styles
    {
        
        public TableOfContentTest(Word.Document doc, Word.Application app)
            : base(doc, app)
        {
           
        }
        public bool runLevels()
        {
            Word.TablesOfContents numOfContents = doc.TablesOfContents;
            if (numOfContents.Count > 0)
            {
                foreach (Word.TableOfContents table in numOfContents)
                {
                    int lowerLevel = table.LowerHeadingLevel;
                    if (lowerLevel > 3)
                    {
                        return false;
                    }
                }
                return true;
            }
            return false;
        }


        public bool runInUse()
        {
            return doc.TablesOfContents.Count != 0;
        }


    }
}
