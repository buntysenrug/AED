using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class Heading1:Styles
    {
        //private Word.Document doc;
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
        private Word.Style heading1;
        
        /*Constructor of Class.
         * 
         */
        public Heading1(Word.Document doc,Word.Application app)
            : base(doc,app)
        {
            //this.doc = doc;
            
            this.spaceBeforeLower = 6f;
            this.spaceBeforeUpper = 30f;
            this.spaceAfterLower = 6f;
            this.spaceAfterUpper = 30f;
            this.outLineAsString = "1";
            this.widow = true;
            this.keepWithNext = true;
            this.quickStyleList = true;
            this.autoUpdate = false;
            this.numbered = true;
            this.bulleted = true;
            foreach (Word.Style s in set)
            {
                if (s.NameLocal.Contains("Heading 1") || s.NameLocal.Equals("Heading 1"))
                {
                    heading1 = s;
                    break;
                }
            }
            
        }

        public bool runInUse()
        {
            try
            {
                return this.heading1.InUse;
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
            if (heading1 != null)
            {
                return style_name.Contains("Heading 1");
            }
            return false;
        }

        /*This method is will run runOutline test on Heading 1 style based as per specifications.
         * 
         */
        public bool runOutline()
        {
            if (heading1 != null)
            {
                return outLineStyleCheck(heading1, Word.WdOutlineLevel.wdOutlineLevel1);
            }
            return false;
        }

        /*A method that will check before spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceB()
        {
            if (heading1 != null)
            {
                return spaceBeforeStyleCheck(heading1, this.spaceBeforeLower, this.spaceBeforeUpper);
            }
            return false;
        }

        /*A method that will check after spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceA()
        {
            if (runInUse())
                return spaceAfterStyleCheck(heading1, this.spaceAfterLower, this.spaceAfterUpper);
            else
                return false;
        }

        /*A method that will run Widow styles check
         * 
         */
        public bool runWidow()
        {
            return widowStyleCheck(heading1, this.widow);
        }

        /*A method that will run runKeep test
         * 
         */
        public bool runKeep()
        {
            if (heading1 != null)
            {
                return keepWithNextStyleCheck(heading1, this.keepWithNext);
            }
            return false;
        }

        /*A method that will checks Automatic Update is off
         * 
         */
        public bool runAUpdate()
        {
            return autoUpdateStyleCheck(this.heading1, this.autoUpdate);
        }

        /*A method that will check for Numbered heading
         * 
         */
        public bool runNumbered()
        {
            if (heading1 != null)
            {
                return numberedStyleCheck(this.heading1, this.numbered);
            }
            return false;
        }

        /*A method that will run check for Bulleted headings
         * 
         */
        public bool runBulleted()
        {
            if (heading1 != null)
            {
                return bulletedStyleCheck(this.heading1, this.bulleted);
            }
            return false;
        }

        /*A method that will run runtotalspace check
         * 
         */
        public bool runTotalSpace()
        {
            if (heading1 != null)
            {
                float total = this.heading1.ParagraphFormat.SpaceBefore + this.heading1.ParagraphFormat.SpaceAfter;
                if (!(total >= 12 && total <= 50))
                {
                    return false;
                }
                return true;
            }
            return false;
        }

        public bool headin1Test()
        {
            if (heading1 != null)
            {
                if (!(heading1.Font.Name.Equals("Times New Roman") || heading1.Font.Name.Equals("Arial")))
                {
                    return false;
                }
                return true;
            }
            return false;
        }
    }
}
