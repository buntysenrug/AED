using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word=Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class ListBulletStyle:Styles
    {
        private Word.Style listbullet;
        private float spaceAfterLower;
        private float spaceAfterUpper;
        private float spaceBeforeLower;
        private float spaceBeforeUpper;

        public ListBulletStyle(Word.Document doc, Word.Application app)
            : base(doc,app)
        {
            foreach (Word.Style current in doc.Styles)
            {
                if (current.NameLocal.Equals("List Bullet"))
                {
                    this.listbullet = current;
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
                if (current.NameLocal.Equals("List Bullet"))
                    return current.InUse;
            }
            return false;
        }

        /*This method will run the runbase test based on Heading 1 Style as per specifications.
         * 
         */
        public bool runBase()
        {
            Word.Style s = getBaseStyle(listbullet.NameLocal);
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
            return outLineStyleCheck(listbullet, Word.WdOutlineLevel.wdOutlineLevelBodyText);
        }

        /*A method that will check before spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceB()
        {
            return spaceBeforeStyleCheck(listbullet, this.spaceBeforeLower, this.spaceBeforeUpper);
        }

        /*A method that will check after spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceA()
        {
            return spaceAfterStyleCheck(listbullet, this.spaceAfterLower, this.spaceAfterUpper);
        }

        public bool runSpaceBetweenPara()
        {
            return listbullet.NoSpaceBetweenParagraphsOfSameStyle;
        }

        /*A method to check the indent of the quote is from left side
         * 
         */
        public bool runIndent()
        {
            if (listbullet.ParagraphFormat.LeftIndent < 0 || listbullet.ParagraphFormat.RightIndent < 0)
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
            float total = this.listbullet.ParagraphFormat.SpaceBefore + this.listbullet.ParagraphFormat.SpaceAfter;
            if (!(total >= 3 && total <= 30))
            {
                return false;
            }
            return true;
        }
    }
}
