using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word=Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class QuoteStyle : Styles
    {
        private Word.Style quote;
        private bool keepWithNext;
        private float spaceAfterLower;
        private float spaceAfterUpper;

        public QuoteStyle(Word.Document doc, Word.Application app)
            : base(doc,app)
        {
            foreach (Word.Style current in doc.Styles)
            {
                if (current.NameLocal.Equals("Quote") || current.NameLocal.Contains("Quote"))
                {
                    this.quote = current;
                    break;
                }
            }
            this.keepWithNext = true;
            this.spaceAfterLower = 6.0f;
            this.spaceAfterUpper = 18.0f;
        }

        public bool runInUse()
        {
            try
            {
                return this.quote.InUse;
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
            Word.Style s = getBaseStyle(quote.NameLocal);
            if (s != null)
            {
                if (s.NameLocal.Equals("Normal"))
                {
                    return true;
                }
            }
            return false;
        }

        /*A method that will run runKeep test
         * 
         */
        public bool runKeep()
        {
            if (this.quote != null)
            {
                return keepWithNextStyleCheck(quote, this.keepWithNext);
            }
            return false;
        }

        /*A method that will check that font style of the quote is italic.
         * 
         */
        public bool runFontStyle()
        {
            if (this.quote != null)
            {
                return this.quote.Font.Italic != 0;
            }
            return false;
        }

        /*A method that will check after spacing in Style and return value according to the specifications.
         * 
         */
        public bool runSpaceA()
        {
            return spaceAfterStyleCheck(quote, this.spaceAfterLower, this.spaceAfterUpper);
        }

        /*A method to check the indent of the quote is from left side
         * 
         */
        public bool runIndent()
        {
            return (quote.ParagraphFormat.LeftIndent >= 0.5 && quote.ParagraphFormat.RightIndent >= 0.5);
        }


    }
}