using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class NormalStyle:Styles
    {
        private Word.Style normalStyle;

        public NormalStyle(String filename):base(filename)
        {
            HashSet<Word.Style> set = getStyles();
            foreach (Word.Style s in set)
            {
                if (s.NameLocal.Equals("Normal"))
                {
                    normalStyle = s;
                    break;
                }
            }
        }

        public Word.Style getNormalStyle()
        {
            return normalStyle;
        }


    }
}
