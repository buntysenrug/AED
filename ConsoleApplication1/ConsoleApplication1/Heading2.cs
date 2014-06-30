using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class Heading2:Styles
    {
        private float spaceBeforeLower;
        private float spaceBeforeUpper;
        private float spaceAfterLower;
        private float spaceAfterUpper;
        private String outLineAsString;
        private bool widow;
        private bool keepWithNext;
        private bool quickStyleList;
        private bool autoUpdate;
        private bool numbered;
        private bool bulleted;
        private Word.Style heading2;
        //private Word.Document doc;

        public Heading2(Word.Document doc, Word.Application app)
            : base(doc,app)
        {
            this.spaceBeforeLower = 6f;
            this.spaceBeforeUpper = 18f;
            this.spaceAfterLower = 6f;
            this.spaceAfterUpper = 18f;
            this.outLineAsString = "2";
            this.widow = true;
            this.keepWithNext = true;
            this.quickStyleList = true;
            this.autoUpdate = false;
            this.numbered = true;
            this.bulleted = true;
           // this.doc = doc;
            foreach (Word.Style s in Styles.set)
            {
                if (s.NameLocal.Equals("Heading 2"))
                {
                    heading2 = s;
                    break;
                }
            }
        }

        public bool runInUse()
        {
            try
            {
                return this.heading2.InUse;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /*This method will run the runbase test based on Heading 1 Style as per specifications.
         * 
         */
        public bool runBase()
        {
            Word.Style s = getBaseStyle(heading2.NameLocal);
            if (s.NameLocal.Equals("Normal") || s.NameLocal.Equals("Heading 1"))
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
            return outLineStyleCheck(heading2, Word.WdOutlineLevel.wdOutlineLevel2);
        }

        /*A method that will check before spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceB()
        {
            return spaceBeforeStyleCheck(heading2, this.spaceBeforeLower, this.spaceBeforeUpper);
        }

        /*A method that will check after spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceA()
        {
            return spaceAfterStyleCheck(heading2, this.spaceAfterLower, this.spaceAfterUpper);
        }

        /*A method that will run Widow styles check
         * 
         */
        public bool runWidow()
        {
            return widowStyleCheck(heading2, this.widow);
        }

        /*A method that will run runKeep test
         * 
         */
        public bool runKeep()
        {
            return keepWithNextStyleCheck(heading2, this.keepWithNext);
        }

        /*A method that will checks Automatic Update is off
         * 
         */
        public bool runAUpdate()
        {
            return autoUpdateStyleCheck(this.heading2, this.autoUpdate);
        }

        /*A method that will check for Numbered heading
         * 
         */
        public bool runNumbered()
        {
            return numberedStyleCheck(this.heading2, this.numbered);
        }

        /*A method that will run check for Bulleted headings
         * 
         */
        public bool runBulleted()
        {
            return bulletedStyleCheck(this.heading2, this.bulleted);
        }

        /*A method that will run runtotalspace check
         * 
         */
        public bool runTotalSpace()
        {
            float total = this.heading2.ParagraphFormat.SpaceBefore + this.heading2.ParagraphFormat.SpaceAfter;
            if (!(total >= 12 && total <= 36))
            {
                return false;
            }
            return true;
        }

        public bool headin2Test()
        {
            if (!(heading2.Font.Name.Equals("Times New Roman") || heading2.Font.Name.Equals("Arial")))
            {
                return false;
            }
            return true;
        }
    }
}
