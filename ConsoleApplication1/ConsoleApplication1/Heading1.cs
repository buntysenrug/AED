﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class Heading1:Styles
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
        private Word.Style heading1;
        
        /*Constructor of Class.
         * 
         */
        public Heading1(Word.Document doc)
            : base(doc)
        {
            this.spaceBeforeLower = 6f;
            this.spaceBeforeUpper = 30f;
            this.spaceAfterLower = 6f;
            this.spaceAfterUpper = 30f;
            this.outLineAsString = "1";
            this.widow = true;
            this.keepWithNext = true;
            this.quickStyleList = true;
            this.autoUpdate = false;
            this.numbered = false;
            this.bulleted = false;
            foreach (Word.Style s in Styles.set)
            {
                if (s.NameLocal.Equals("Heading 1"))
                {
                    heading1 = s;
                    break;
                }
            }
            
        }

        public bool runInUse()
        {
            return Styles.style_name.Contains("Heading 1");
        }

        /*This method will run the runbase test based on Heading 1 Style as per specifications.
         * 
         */
        public bool runBase()
        {
            Word.Style s = getBaseStyle(heading1.NameLocal);
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
            return outLineStyleCheck(heading1, Word.WdOutlineLevel.wdOutlineLevel1);
        }

        /*A method that will check before spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceB()
        {
            return spaceBeforeStyleCheck(heading1, this.spaceBeforeLower, this.spaceBeforeUpper);
        }

        /*A method that will check after spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceA()
        {
            return spaceAfterStyleCheck(heading1, this.spaceAfterLower, this.spaceAfterUpper);
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
            return keepWithNextStyleCheck(heading1, this.keepWithNext);
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
            return numberedStyleCheck(this.heading1, this.numbered);
        }

        /*A method that will run check for Bulleted headings
         * 
         */
        public bool runBulleted()
        {
            return bulletedStyleCheck(this.heading1, this.bulleted);
        }

        /*A method that will run runtotalspace check
         * 
         */
        public bool runTotalSpace()
        {
            float total = this.heading1.ParagraphFormat.SpaceBefore + this.heading1.ParagraphFormat.SpaceAfter;
            if (!(total >= 12 && total <= 50))
            {
                return false;
            }
            return true;
        }
    }
}