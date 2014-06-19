using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class Heading5:Heading3
    {
        private Word.Style heading5;

        public Heading5(Word.Document doc)
            : base(doc)
        {
            foreach (Word.Style s in Styles.set)
            {
                if (s.NameLocal.Equals("Heading 5"))
                {
                    heading5 = s;
                    break;
                }
            }
        }

        /*This method will run the runbase test based on Heading 5 Style as per specifications.
         * 
         */
        public override bool runBase()
        {
            Word.Style s = getBaseStyle(heading5.NameLocal);
            if (s.NameLocal.Equals("Normal") || s.NameLocal.Equals("Heading 1") || s.NameLocal.Equals("Heading 2") || s.NameLocal.Equals("Heading 3") || s.NameLocal.Equals("Heading 4"))
            {
                return true;
            }
            return false;
        }

        /*A method that checks whether Heading 5 is used or not.
         * 
         */
        public override bool runInUse()
        {
            return Styles.style_name.Contains("Heading 5");
        }

        /*This method is will run runOutline test on Heading 5 style based as per specifications.
         * 
         */
        public override bool runOutline()
        {
            return outLineStyleCheck(heading5, Word.WdOutlineLevel.wdOutlineLevel5);
        }
    }
}
