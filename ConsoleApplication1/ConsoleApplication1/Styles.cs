using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class Styles
    {
        private Word.Application app;
        private Word.Document doc=null;
        private HashSet<Word.Style> set;
        
       
        /*Constructor of the Base Class Styles
         * This constructor takes in the filename of word file as parameter
         */
        public Styles(String filename)
        {
            app = new Word.Application();
            doc = app.Documents.Open(filename);
            app.Visible=false;
            this.set = new HashSet<Word.Style>();
            getStyles();    
        }

        /*This method retuns a Hashset of all the styles that are used in the word document
         * In Simple words it returns a set of all Styles used in document.
         */
        public HashSet<Word.Style> getStyles()
        {
            foreach (Word.Style s in doc.Styles)
            {
                this.set.Add(s);
            }
            return this.set;
        }

        /*This method prints all the styles that were in the hash set of styles
         *Warning :- This method should be called after getStyles() Method as the set would be empty if
         * called earlier.
         */
        public void printStyles()
        {
            foreach (Word.Style s in this.set)
            {
                Console.WriteLine(s.NameLocal);
            }
        }


        /*This method returns a Dictionary of type <Style, Based on Style>
         * This method has a set of two styles, indicating that second one is based on first one.
         */
        public Dictionary<Word.Style, Word.Style> getBaseStyles()
        {
            Dictionary<Word.Style, Word.Style> dictionary_base_styles = new Dictionary<Word.Style, Word.Style>();
            foreach (Word.Style s in doc.Styles)
            {
                dictionary_base_styles.Add(s, s.get_BaseStyle());
            }
            return dictionary_base_styles;
        }

        /*This method will return a dictionary of type <Style, Fontsize> used in that style
         * This dictionary can be used to refer the font size of a particular style
         */
        public Dictionary<Word.Style, Double> getFontSizeByStyles()
        {
            Dictionary<Word.Style, Double> dictionary_font_size_by_styles = new Dictionary<Word.Style, double>();
            foreach (Word.Style s in this.set)
            {
                dictionary_font_size_by_styles.Add(s, s.Font.Size);
            }

            return dictionary_font_size_by_styles;
        }

        /*This method returns a dictionary of type <Style, bool> where bool represents
         * true or false based on the fact that bold italic or underline is used.
         */
        public Dictionary<Word.Style, bool> fontStyleUsed()
        {
            Dictionary<Word.Style,bool> dictionary_font_style_used = new Dictionary<Word.Style, bool>();
            foreach (Word.Style s in this.set)
            {
                bool ans=false;
                if (s.Font.Bold != 0 || s.Font.Italic != 0 || s.Font.Underline != 0)
                {
                    ans = true;
                }
                dictionary_font_style_used.Add(s, ans);
            }
            return dictionary_font_style_used;
        }


        /*This method will return bool value based on type Style where bool represents true or false based 
         * on the fact whether style is widow or not 
         */
        public bool widowStyleCheck(Word.Style style, bool widow)
        {
            bool styleOK = true;
            int widowNum = 0;
            if (widow)
            {
                widowNum = -1;
            }
            if (style.ParagraphFormat.WidowControl != widowNum)
            {
                styleOK = false;
            }

            return styleOK;
        }


        /*This method will return the before space based on type Style
         * where spacing before is a floating point value.
         */
        public float getSpacingBefore(Word.Style style)
        {
            return style.ParagraphFormat.SpaceBefore;
        }

        /*This method will return After space based on Style
         * where spacing After is a floating point value.
         */
        public float getSpacingAfter(Word.Style style)
        {
            return style.ParagraphFormat.SpaceAfter;
        }

        /*A method that will check keep with style with the style passed in.
         * 
         */
        public bool keepWithNextStyleCheck(Word.Style style, bool keepWithNext)
        {
            bool styleOK = true;
            int keepWithNextNum = 0;
            if (keepWithNext)
            {
                keepWithNextNum = -1;
            }

            if (style.ParagraphFormat.KeepWithNext != keepWithNextNum)
            {
                styleOK = false;
            }
            return styleOK;
        }

        /*
        * A method that will check auto update of a style
        * */
        public bool autoUpdateStyleCheck(Word.Style style, bool autoUpdate)
        {
            bool styleOK = true;
            if (style.AutomaticallyUpdate != autoUpdate)
            {
                styleOK = false;
            }
            return styleOK;
        }

        /*
      * A method that will check if style is numbered
      * */
        public bool numberedStyleCheck(Word.Style style, bool numbering)
        {//check for numbering of bullets     
            String[] description = style.Description.Split(' ');
            for (int i = 0; i < description.Length; i++)
            {
                if (description[i].Equals("Numbered", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (!numbering)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            if (numbering)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /*A method that will check whether style is bulleted or not.
         * 
         */
        public bool bulletedStyleCheck(Word.Style style, bool bullets)
        {//check for numbering of bullets 
            String[] description = style.Description.Split(' ');
            for (int i = 0; i < description.Length; i++)
            {
                if (description[i].Equals("Bulleted"))
                {
                    if (!bullets)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            if (bullets)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        


    }
}
