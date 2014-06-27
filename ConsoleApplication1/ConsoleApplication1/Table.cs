using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class Table:Styles
    {
        private double leftMargin;
        private double rightMargin;
        private double numOfTableCaps;

        public Table(Word.Document doc, Word.Application app,double left_margin,double right_margin,int number_of_tablecaps)
            : base(doc, app)
        {
            this.leftMargin = left_margin;
            this.rightMargin = right_margin;
            this.numOfTableCaps = number_of_tablecaps;
        }

        /*
         * A method that will check the tables in the doc
         */
        public bool tableTests(bool runTextWrap, bool runWidth, bool runTableCount, bool runCaption, bool runTablePos, bool runTableStyle, bool runTablePsy)
        {

            bool tableOK = true;
            //FeedbackInfo feed = null;
            //float tableTemp = 4 * numberOfTables;

            foreach (Word.Table table in doc.Tables)
            {
                String tableTitle = table.Title;

                if (table.Rows.WrapAroundText == -1)
                {
                    if (runTextWrap)
                    {
                        tableOK = false;
                        
                        runTextWrap = false;
                    }
                }

                float width = 0;
                width = app.PointsToCentimeters(width);
                float indent = table.Rows.LeftIndent;
                indent = app.PointsToCentimeters(indent);
                double a4Width = 21;
                double total = a4Width - leftMargin - rightMargin - indent;

                Word.Range tableRange = table.Range;

                foreach (Word.Cell cell in tableRange.Cells)
                {
                    width += cell.Width;
                }
                width = width / (table.Rows.Count);

                width = app.PointsToCentimeters(width);
                if (!(width <= total + 1))
                {
                    if (runWidth)
                    {
                        tableOK = false;
                        runWidth = false;
                    }

                }

                if (!tableHasCaption(table))
                {
                    if (runCaption)
                    {
                        tableOK = false;
                        runCaption = false;
                    }
                }

                if (doc.Tables.Count > numOfTableCaps)
                {
                    if (runTableCount)
                    {
                        tableOK = false;
                        runTableCount = false;//stop test now as found one
                    }
                }

                if (!(table.Rows.Alignment == Word.WdRowAlignment.wdAlignRowCenter))
                {
                    if (runTablePos)
                    {
                        tableOK = false;
                        runTablePos = false;//stop test
                    }
                }

                Word.Style thestyle = table.get_Style();
                bool styleTest = false;
                Word.Range tabRange = table.Range;

                Word.Rows rows = table.Rows;

                foreach (Word.Cell cell in tabRange.Cells)
                {
                    Word.Style rowStyle = cell.Range.get_Style();

                    if (rowStyle.NameLocal.Equals("Table Header") || rowStyle.NameLocal.Equals("Table Body"))
                    {
                        styleTest = true;
                        runTableStyle = false;
                        break;
                    }
                }

                if (thestyle.NameLocal.Equals("Table Grid") && !styleTest)
                {
                    if (runTableStyle)
                    {
                        tableOK = false;
                        runTableStyle = false;
                    }
                }

                Word.Style thestyle2 = table.get_Style();

                Word.Range tabRange2 = table.Range;

                Word.Rows rows2 = table.Rows;

                foreach (Word.Cell cell in tabRange2.Cells)
                {
                    Word.Style rowStyle = cell.Range.get_Style();

                    if ((!rowStyle.NameLocal.Equals("TablePsychology") && !rowStyle.NameLocal.Equals("TableLargePsychology")) && (!cell.Range.Text.Equals("") && !cell.Range.Text.Equals("\r\a")))
                    {
                        if (runTablePsy)
                        {
                            tableOK = false;
                            runTablePsy = false;
                        }
                    }
                }

            }

            return tableOK;
        }

        //test if an image has a caption preceeding it 
        private bool tableHasCaption(Word.Table table)
        {
            Word.Paragraph p = table.Range.Paragraphs.First;
            Word.Paragraph paraBefore = p.Previous();
            if (paraBefore != null)
            {
                foreach (Word.Field f in paraBefore.Range.Fields)
                {
                    if (f.Type == Word.WdFieldType.wdFieldSequence)
                    {
                        return true;
                    }
                }
            }

            if (paraBefore != null)
            {
                Word.Paragraph paraBefore2 = paraBefore.Previous();
                if (paraBefore2 != null)
                {
                    foreach (Word.Field f in paraBefore2.Range.Fields)
                    {
                        if (f.Type == Word.WdFieldType.wdFieldSequence)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
    }

}
