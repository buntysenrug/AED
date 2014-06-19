using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class Heading3:Styles
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
        private Word.Style heading3;

        public Heading3(Word.Document doc)
            : base(doc)
        {
            this.spaceBeforeLower = 6f;
            this.spaceBeforeUpper = 12f;
            this.spaceAfterLower = 6f;
            this.spaceAfterUpper = 12f;
            this.outLineAsString = "3";
            this.widow = true;
            this.keepWithNext = true;
            this.quickStyleList = true;
            this.autoUpdate = false;
            this.numbered = false;
            this.bulleted = false;
            foreach (Word.Style s in Styles.set)
            {
                if (s.NameLocal.Equals("Heading 3"))
                {
                    heading3 = s;
                    break;
                }
            }
        }

        public bool runInUse()
        {
            return Styles.style_name.Contains("Heading 2");
        }

        /*This method will run the runbase test based on Heading 1 Style as per specifications.
         * 
         */
        public bool runBase()
        {
            Word.Style s = getBaseStyle(heading3.NameLocal);
            if (s.NameLocal.Equals("Normal") || s.NameLocal.Equals("Heading 1") || s.NameLocal.Equals("Heading 2"))
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
            return outLineStyleCheck(heading3, Word.WdOutlineLevel.wdOutlineLevel3);
        }

        /*A method that will check before spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceB()
        {
            return spaceBeforeStyleCheck(heading3, this.spaceBeforeLower, this.spaceBeforeUpper);
        }

        /*A method that will check after spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceA()
        {
            return spaceAfterStyleCheck(heading3, this.spaceAfterLower, this.spaceAfterUpper);
        }

        /*A method that will run Widow styles check
         * 
         */
        public bool runWidow()
        {
            return widowStyleCheck(heading3, this.widow);
        }

        /*A method that will run runKeep test
         * 
         */
        public bool runKeep()
        {
            return keepWithNextStyleCheck(heading3, this.keepWithNext);
        }

        /*A method that will checks Automatic Update is off
         * 
         */
        public bool runAUpdate()
        {
            return autoUpdateStyleCheck(this.heading3, this.autoUpdate);
        }

        /*A method that will check for Numbered heading
         * 
         */
        public bool runNumbered()
        {
            return numberedStyleCheck(this.heading3, this.numbered);
        }

        /*A method that will run check for Bulleted headings
         * 
         */
        public bool runBulleted()
        {
            return bulletedStyleCheck(this.heading3, this.bulleted);
        }

        /*A method that will run runtotalspace check
         * 
         */
        public bool runTotalSpace()
        {
            float total = this.heading3.ParagraphFormat.SpaceBefore + this.heading3.ParagraphFormat.SpaceAfter;
            if (!(total >= 12 && total <= 26))
            {
                return false;
            }
            return true;
        }
    }
   }

