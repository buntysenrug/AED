using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class TitleStyle:Styles
    {
        private Word.Style title;
        public TitleStyle(Word.Document doc)
            : base(doc)
        {
            foreach (Word.Style s in Styles.set)
            {
                if (s.NameLocal.Contains("Title") || s.NameLocal.Equals("Title"))
                {
                    this.title = s;
                    break;
                }
            }

        }

        public bool runTitleUsed()
        {
            try
            {
                return this.title.InUse;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool runTitleNotTwice(List<Word.Style> stylesindoc,int titlecount)
        {
            foreach (Word.Style s in stylesindoc)
            {
                if (s.NameLocal.Equals("Heading 1") || s.NameLocal.Contains("Heading 1"))
                {
                    if (titlecount > 1)
                    {
                        return false;
                    }
                    
                }
            }
            return true;
        }

        public bool runTitleThreeTimes(bool titleusedcount)
        {
            return titleusedcount;
        }


        
    }
}
