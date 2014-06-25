using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class TableOfContentStyle:Styles
    {
        private Word.Style tocstyle;
        public TableOfContentStyle(Word.Document doc)
            : base(doc)
        {
            foreach (Word.Style s in set)
            {
                if (s.NameLocal.Equals("Table of Content") || s.NameLocal.Contains("TOC"))
                {
                    tocstyle = s;
                    break;
                }
            }

        }

        public bool runBase()
        {
            Word.Style base_style=tocstyle.get_BaseStyle();
            if (base_style.NameLocal.Equals("Normal"))
            {
                return true;
            }
            return false;
        }

        public bool runOutline()
        {
            return outLineStyleCheck(tocstyle, Word.WdOutlineLevel.wdOutlineLevelBodyText);
        }

        public bool runKeep()
        {
            return keepWithNextStyleCheck(tocstyle, false);
        }

    }
}
