﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class NormalStyle : Styles
    {
       
        private Word.Style normalStyle;
        private float fontSizeLower;
        private float fontSizeUpper;
        private bool keepLinesTogether;
        private bool pageBreakBefore;
        private bool keepWithNext;
        private int keepTogetherNum;
        private int pageBreakNum;

        /*Initilization of base class constructor and also this class
         * known as Derived class i.e. NormalStyle
         */
        public NormalStyle(Word.Document doc,Word.Application app)
            : base(doc,app)
        {
            //HashSet<Word.Style> set = getStyles(doc);
            foreach (Word.Style s in set)
            {
                if (s.NameLocal.Equals("Normal") || s.NameLocal.Contains("Normal"))
                {
                    normalStyle = s;
                    break;
                }
            }
            this.fontSizeLower = 11f;
            this.fontSizeUpper = 13f;
            this.keepLinesTogether = false;
            this.pageBreakBefore = false;
            this.keepWithNext = false;
            this.keepTogetherNum = 0;
            this.pageBreakNum = 0;
            if (pageBreakBefore)
            {
                pageBreakNum = -1;
            }
            if (keepLinesTogether)
            {
                keepTogetherNum = -1;
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
            if (normalStyle != null)
            {
                Word.Style s = getBaseStyle(normalStyle.NameLocal);
                if (s.NameLocal.Equals(""))
                {
                    return true;
                }
                return false;
            }
            return false;
        }

        /*This method will runfontsize test on normal style based on as per specifications.
         * 
         */
        public bool runFontSize()
        {
            //Dictionary<Word.Style, Double> dict = getFontSizeByStyles();
            //Double size = dict[normalStyle];
            if (normalStyle != null)
            {
                if (10 < normalStyle.Font.Size && normalStyle.Font.Size <= 12)
                {
                    return true;
                }
                return false;
            }
            return false;
        }

        /*This method is will run runOutline test on normal style based as per specifications.
         * 
         */
        public bool runOutline()
        {
            if (normalStyle != null)
            {
                return outLineStyleCheck(normalStyle, Word.WdOutlineLevel.wdOutlineLevelBodyText);
            }
            return false;
        }

        /*A method that will check after spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceA()
        {
            if (normalStyle != null)
            {
                return spaceAfterStyleCheck(normalStyle, 12);
            }
            return false;
        }

        /*A method that will check before spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceB()
        {
            if (normalStyle != null)
            {
                return spaceBeforeStyleCheck(normalStyle, 0);
            }
            return false;
        }

        /*A method that will run runKeep test
         * 
         */
        public bool runKeep()
        {
            if (normalStyle != null)
            {
                return keepWithNextStyleCheck(normalStyle, keepWithNext);
            }
            return false;
        }

        /*A method that will run runtotalspace check
         * 
         */
        public bool runTotalSpace()
        {
            if (normalStyle != null)
            {
                float total = normalStyle.ParagraphFormat.SpaceBefore + normalStyle.ParagraphFormat.SpaceAfter;
                if (!(total >= 3 && total <= 30))
                {
                    return false;
                }
                return true;
            }
            return false;
        }

        /*A method that will run linespacing check.
         * 
         */
        public bool runLineSpacing(Word.Application app)
        {
            if (normalStyle != null)
            {
                float lines = app.PointsToLines(normalStyle.ParagraphFormat.LineSpacing);
                if (lines != 1.5f || lines != 2.0f || lines != 3.0f)
                {
                    return false;
                }
                return true;
            }
            return false;
        }

        /*A method that will run FontStyle check.
         * 
         */
        public bool runFontStyle()
        {
            if (normalStyle != null)
            {
                if (normalStyle.Font.Bold != 0 && normalStyle.Font.Italic != 0 &&
                    normalStyle.Font.Underline != 0 && normalStyle.Font.ItalicBi != 0)
                {
                    return false;
                }
                return true;
            }
            return false;
        }

        /*A method that will run Font effects Test
         * 
         */
        public bool runFontEffects()
        {
            if (normalStyle != null)
            {
                if (normalStyle.Font.StrikeThrough != 0 || normalStyle.Font.DoubleStrikeThrough != 0 ||
                    normalStyle.Font.Superscript != 0 || normalStyle.Font.Subscript != 0 ||
                    normalStyle.Font.SmallCaps != 0 || normalStyle.Font.AllCaps != 0 || normalStyle.Font.Hidden != 0)
                {
                    return false;
                }
                return true;
            }
            return false;
        }

        /*A method that will run Lines Together test
         * 
         */
        public bool runLinesTogether()
        {
            if (normalStyle != null)
            {
                if (normalStyle.ParagraphFormat.KeepTogether != keepTogetherNum)
                {
                    return false;
                }
                return true;
            }
            return false;
        }

        /*A method that will run page break test
         * 
         */
        public bool runPageBreak()
        {
            if (normalStyle != null)
            {
                if (normalStyle.ParagraphFormat.PageBreakBefore != pageBreakNum)
                {
                    return false;
                }
                return true;
            }
            return false;
        }

        /*A method that will run Widow styles check
         * 
         */
        public bool runWidow()
        {
            if (normalStyle != null)
            {
                return widowStyleCheck(normalStyle, true);
            }
            return false;
        }

        public bool runInUse()
        {
            foreach (Word.Style current in doc.Styles)
            {
                if (current.NameLocal.Equals("Normal"))
                    return current.InUse;
            }
            return false;
        }

        public bool normalTest()
        {
            if (normalStyle != null)
            {
                if (!(normalStyle.Font.Name.Equals("Times New Roman") || normalStyle.Font.Name.Equals("Arial")))
                {
                    return false;
                }
                return true;
            }
            return false;
        }
    }
}
