using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word=Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class CaptionStyle : Styles
    {
       // private Word.Document doc;
        private Word.Style caption;
        private float spaceAfterLower;
        private float spaceAfterUpper;
        private bool widow;
        private bool keepWithNext;
        private int numberOfImages;
        private int numOfTableCaps;
        private int numOfFigureCaps;
        private bool numbered;
        private bool bulleted;

        public CaptionStyle(Word.Document doc)
            : base(doc)
        {
           // this.doc = doc;
            this.spaceAfterLower = 6f;
            this.spaceAfterUpper = 18f;
            this.widow = true;
            this.keepWithNext = true;
            this.numOfFigureCaps = 0;
            this.numOfTableCaps = 0;
            this.numberOfImages = 0;
            this.numbered = true;
            this.bulleted = true;
            foreach (Word.Style current in doc.Styles)
            {
                if (current.NameLocal.Equals("Caption"))
                {
                    this.caption = current;
                    break;
                }
            }
        }
        /*A method that will check that captions are used only if images or tables are there and also number
         * of captions is equal to number of shapes/figures
         */
        public bool captionNoObjects()
        {
            foreach (Word.InlineShape shape in doc.InlineShapes)
            {
                if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture || shape.Type == Word.WdInlineShapeType.wdInlineShapeChart)
                {
                    numberOfImages++;
                }
            }
            foreach (Word.Field f in doc.Fields)
            {
                if (f.Type == Word.WdFieldType.wdFieldSequence)
                {
                    Word.Range range = f.Code;
                    Word.Style captionStyle = range.get_Style();

                    String capString = range.Text;

                    bool isMatchTable = System.Text.RegularExpressions.Regex.IsMatch(capString.ToLower(), "table");
                    bool isMatchFigure = System.Text.RegularExpressions.Regex.IsMatch(capString.ToLower(), "figure");

                    if (isMatchTable)
                    {
                        numOfTableCaps++;
                    }
                    else if (isMatchFigure)
                    {
                        numOfFigureCaps++;
                    }
                }
            }
            if (numberOfImages == numOfFigureCaps && numOfTableCaps == doc.Tables.Count)
            {
                return true;
            }
            return false;
        }

        /*This method will run the runbase test based on Heading 1 Style as per specifications.
         * 
         */
        public bool runBase()
        {
            Word.Style s = getBaseStyle(caption.NameLocal);
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
            return outLineStyleCheck(caption, Word.WdOutlineLevel.wdOutlineLevelBodyText);
        }

        /*A method that will check after spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceA()
        {
            return spaceAfterStyleCheck(caption, this.spaceAfterLower, this.spaceAfterUpper);
        }

        /*A method that will run Widow styles check
         * 
         */
        public bool runWidow()
        {
            return widowStyleCheck(caption, this.widow);
        }

        /*A method that will run runKeep test
         * 
         */
        public bool runKeep()
        {
            return keepWithNextStyleCheck(caption, this.keepWithNext);
        }

        public bool runPosition()
        {
            return true;
        }

        /*A method that will check for Numbered heading
         * 
         */
        public bool runNumbered()
        {
            return numberedStyleCheck(this.caption, this.numbered);
        }

        /*A method that will run check for Bulleted headings
         * 
         */
        public bool runBulleted()
        {
            return bulletedStyleCheck(this.caption, this.bulleted);
        }


        public bool runInUse()
        {
            return ((this.caption.InUse) && captionNoObjects());
        }

        /*A method that will run runtotalspace check
         * 
         */
        public bool runTotalSpace()
        {
            float total = this.caption.ParagraphFormat.SpaceBefore + this.caption.ParagraphFormat.SpaceAfter;
            if (!(total >= 3 && total <= 30))
            {
                return false;
            }
            return true;
        }
    }
}
            
        
