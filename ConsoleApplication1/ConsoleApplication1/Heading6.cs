using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class Heading6:Heading3
    {private Word.Style heading6;

        public Heading6(Word.Document doc)
            : base(doc)
        {
            foreach (Word.Style s in Styles.set)
            {
                if (s.NameLocal.Equals("Heading 6"))
                {
                    heading6 = s;
                    break;
                }
            }
        }

        /*This method will run the runbase test based on Heading 6 Style as per specifications.
         * 
         */
        public override bool runBase()
        {
            Word.Style s = getBaseStyle(heading6.NameLocal);
            if (s.NameLocal.Equals("Normal") || s.NameLocal.Equals("Heading 1") || s.NameLocal.Equals("Heading 2") || s.NameLocal.Equals("Heading 3") || s.NameLocal.Equals("Heading 4") || s.NameLocal.Equals("Heading 5"))
            {
                return true;
            }
            return false;
        }

        /*A method that checks whether Heading 6 is used or not.
         * 
         */
        public override bool runInUse()
        {
            return Styles.style_name.Contains("Heading 6");
        }

        /*This method is will run runOutline test on Heading 6 style based as per specifications.
         * 
         */
        public override bool runOutline()
        {
            return outLineStyleCheck(heading6, Word.WdOutlineLevel.wdOutlineLevel6);
        }
    }
}
