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

        /*Initilization of base class constructor and also this class
         * known as Derived class i.e. NormalStyle
         */
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

        
        /*This method returns the instance variable of Style that stores the normal Style
         * 
         */
        public Word.Style getNormalStyle()
        {
            return normalStyle;
        }

        /*This method will run the runbase test based on Normal Style as per specifications.
         * 
         */
        public bool runBase()
        {
            Word.Style s = getBaseStyle(normalStyle.NameLocal);
            if (s.NameLocal.Equals(""))
            {
                return true;
            }
            return false;
        }

        /*This method will runfontsize test on normal style based on as per specifications.
         * 
         */
        public bool runFontSize()
        {
            Dictionary<Word.Style, Double> dict = getFontSizeByStyles();
            Double size = dict[normalStyle];
            if (10 < size && size <= 12)
            {
                return true;
            }
            return false;
        }

        /*This method is will run runOutline test on normal style based as per specifications.
         * 
         */
        public bool runOutline()
        {
            return outLineStyleCheck(normalStyle, Word.WdOutlineLevel.wdOutlineLevelBodyText);
        }

    }
}
