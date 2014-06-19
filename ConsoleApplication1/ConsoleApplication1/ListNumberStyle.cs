using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word=Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class ListNumberStyle:Styles
    {
        private Word.Style listnumber;
        private float spaceAfterLower;
        private float spaceAfterUpper;
        private float spaceBeforeLower;
        private float spaceBeforeUpper;

        public ListNumberStyle(Word.Document doc)
            :base(doc){
                foreach (Word.Style current in doc.Styles)
                {
                    if (current.NameLocal.Equals("List Number"))
                    {
                        this.listnumber = current;
                        break;
                    }
                }
                this.spaceAfterLower = 0.0f;
                this.spaceAfterUpper = 12.0f;
                this.spaceBeforeLower = 0.0f;
                this.spaceBeforeUpper = 12.0f;
        }

        public bool runInUse()
        {
            foreach (Word.Style current in doc.Styles)
            {
                if (current.NameLocal.Equals("List Number"))
                    return current.InUse;
            }
            return false;
        }

        /*This method will run the runbase test based on Heading 1 Style as per specifications.
         * 
         */
        public bool runBase()
        {
            Word.Style s = getBaseStyle(listnumber.NameLocal);
            if (s.NameLocal.Equals("Normal"))
            {
                return true;
            }
            return false;
        }

        /*This method is will run runOutline test on Heading 1 style based as per specifications.
         * 
         */
        public bool runOutline()
        {
            return outLineStyleCheck(listnumber, Word.WdOutlineLevel.wdOutlineLevelBodyText);
        }

        /*A method that will check before spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceB()
        {
            return spaceBeforeStyleCheck(listnumber, this.spaceBeforeLower, this.spaceBeforeUpper);
        }

        /*A method that will check after spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceA()
        {
            return spaceAfterStyleCheck(listnumber, this.spaceAfterLower, this.spaceAfterUpper);
        }

        public bool runSpaceBetweenPara()
        {
            return listnumber.NoSpaceBetweenParagraphsOfSameStyle;
        }

        /*A method to check the indent of the quote is from left side
         * 
         */
        public bool runIndent()
        {
            if (listnumber.ParagraphFormat.LeftIndent < 0 || listnumber.ParagraphFormat.RightIndent < 0)
            {
                return false;
            }
            return true;
        }

        /*A method that will run runtotalspace check
         * 
         */
        public bool runTotalSpace()
        {
            float total = this.listnumber.ParagraphFormat.SpaceBefore + this.listnumber.ParagraphFormat.SpaceAfter;
            if (!(total >= 3 && total <= 30))
            {
                return false;
            }
            return true;
        }
    }
}
