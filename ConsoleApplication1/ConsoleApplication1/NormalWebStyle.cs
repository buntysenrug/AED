using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class NormalWebStyle:Styles
    {
        private Word.Style normalWeb;
        public NormalWebStyle(Word.Document doc,Word.Application app)
            : base(doc,app)
        {
            foreach (Word.Style s in set)
            {
                if (s.NameLocal.Equals("Normal"))
                {
                    normalWeb = s;
                    break;
                }
            }
        }

        public bool normalWebStyleUsedTest(List<String> normalwebquotes)
        {
            if (normalwebquotes.Count > 0)
            {
                return false;
            }
            return true;
        }

        public bool normalWebStyleUsedTest()
        {
            if (normalWeb != null)
            {
                return !normalWeb.InUse;
            }
            return true;
        }


    }
}
