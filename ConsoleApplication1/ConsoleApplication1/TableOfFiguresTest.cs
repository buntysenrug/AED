using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class TableOfFiguresTest:Styles
    {
        public TableOfFiguresTest(Word.Document doc, Word.Application app)
            : base(doc, app)
        {

        }

        public bool tableOfFiguresTest()
        {
            if (!(doc.TablesOfFigures.Count > 0))
            {
               /* if (numberOfImages > 0 || numberOfTables > 0)
                {
                   
                }*/
                return false;
            }
            return true;
        }
    }
}
