using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class TableOfFigureStyle:Styles
    {
        public TableOfFigureStyle(Word.Document doc):base(doc)
        {

        }

        
    }
}
