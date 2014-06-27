using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class CaptionRefTest:Styles
    {
        private int numOfFigureCaps;
        private int numOfTableCaps;
        private int numberOfImages;
        private int numberOfTables;
        private List<string> noExplanCaptionText;
        private List<string> captionWithObjects;
        private List<string> allCaptionText;

        public CaptionRefTest(Word.Document doc, Word.Application app,int number_of_figurecaps,
            int number_of_tabcaps,int number_of_Img,List<string> no_explain_cap,List<string> all_caption,List<string> cap_with_obj)
            : base(doc, app)
        {
            this.numOfFigureCaps = number_of_figurecaps;
            this.numOfTableCaps = number_of_tabcaps;
            this.noExplanCaptionText = no_explain_cap;
            this.captionWithObjects = cap_with_obj;
            this.allCaptionText = all_caption;
            this.numberOfImages = number_of_Img;
            this.numberOfTables = doc.Tables.Count;
        }

        public bool runObjectCount()
        {
            if (numOfFigureCaps + numOfTableCaps != numberOfImages + numberOfTables)
            {
                return false;
            }
            return true;
        }

        public bool runTextBox()
        {
            foreach (Word.Shape s in doc.Shapes)
            {
                if (s.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                {
                    Word.Range range = s.TextFrame.ContainingRange;
                    Word.Style theStyle = range.get_Style();
                    foreach (Word.Field field in range.Fields)
                    {
                        if (field.Type == Word.WdFieldType.wdFieldSequence)
                        {
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        public bool runExplanation()
        {
            if (noExplanCaptionText.Count > 0)
            {
                return false;
            }
            return true;
        }

        public bool runChecklabel()
        {
            foreach (String s in allCaptionText)
            {
                bool isMatch = System.Text.RegularExpressions.Regex.IsMatch(s.ToLower(), "(?:figure|table)\\s+\\d+.{1}.*");
                if (!isMatch)
                {
                    return false;
                }
            }
            return true;
        }


    }
}
