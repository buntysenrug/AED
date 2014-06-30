using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class CrossRefTest : Styles
    {

        private int number_of_images;
        private int number_of_figure_caps;
        private int number_of_table_caps;
        private int number_of_table;
        private int refLink;
        private int numOfCrossRef;
        private bool preliminary_check;

        public CrossRefTest(Word.Document doc, Word.Application app, int numberOfTableCaps, int numOfFigureCaps,
            int numberOfImages, int numberOfTables, int crossRefLink, int crossRefNum)
            : base(doc, app)
        {
            this.number_of_table_caps = numberOfTableCaps;
            this.number_of_images = numberOfImages;
            this.number_of_table = numberOfTables;
            this.number_of_table_caps = numberOfTableCaps;
            this.refLink = crossRefLink;
            this.numOfCrossRef = crossRefNum;
            this.preliminary_check = !(number_of_figure_caps > 0 || number_of_table_caps > 0) &&
                (number_of_images > 0 || number_of_table > 0);
        }

        public bool runLink()
        {
            if (this.refLink > 0)
            {
                return false;
            }
            return true;
        }

        public bool runNumOfRef()
        {
            object tableCross = doc.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionTable);
            Array arr = ((Array)(tableCross));

            object figureCross = doc.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionFigure);
            Array arrFigure = ((Array)(figureCross));

            String[] allCross = new String[arr.Length + arrFigure.Length];

            arr.CopyTo(allCross, 0);
            arrFigure.CopyTo(allCross, arr.Length);
            if (!(this.numOfCrossRef >= number_of_figure_caps + number_of_table_caps))
            {
                foreach (Word.Field field in doc.Fields)
                {
                    String s = field.Code.Text;
                    Word.Range result = field.Result;
                    if (field.Type == Word.WdFieldType.wdFieldRef)
                    {
                        bool foundMatch = false;
                        for (int i = 0; i < allCross.Length; i++)
                        {
                            string myHeading = (string)allCross.GetValue(i);
                            if (myHeading != null)
                            {
                                String range = field.Result.Text;
                                bool isMatch = System.Text.RegularExpressions.Regex.IsMatch(myHeading, range);
                                if (isMatch)
                                {
                                    foundMatch = true;
                                    allCross.SetValue(null, i);
                                    break;
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < allCross.Length; i++)
                {
                    if (allCross.GetValue(i) != null)
                    {
                        return false;
                    }
                }
            }
            return true;

        }
    }
}
