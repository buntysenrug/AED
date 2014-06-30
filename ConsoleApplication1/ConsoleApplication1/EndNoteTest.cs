using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class EndNoteTest:Styles
    {
        private bool runSub;
        private bool intext;
        private bool biblio;
        public EndNoteTest(Word.Document doc, Word.Application app)
            : base(doc, app)
        {
            foreach (Word.Field field in doc.Fields)
            {
                String s = field.Code.Text;
                if (field.Type == Word.WdFieldType.wdFieldAddin)
                {
                    String text = field.Code.Text;
                    String[] intextSplit = new String[1];
                    intextSplit[0] = "EndNote";
                    String[] result = text.Split(intextSplit, StringSplitOptions.None);
                    if (result.Length > 1)
                    {
                        runSub = true;
                        break;
                    }
                    else
                    {
                        String[] biblioSplit = new String[1];
                        biblioSplit[0] = "REFLIST";
                        if (text.Split(biblioSplit, StringSplitOptions.None).Length > 1)
                        {
                            runSub = true;
                            break;
                        }
                    }
                }
            }

            if (runSub)
            {
                foreach (Word.Field field in doc.Fields)
                {
                    String s = field.Code.Text;
                    if (field.Type == Word.WdFieldType.wdFieldAddin)
                    {
                        //if (runIntext)
                        
                            String text = field.Code.Text;
                            String[] intextSplit = new String[1];
                            intextSplit[0] = "EndNote";
                            String[] result = text.Split(intextSplit, StringSplitOptions.None);
                            if (result.Length > 1)
                            {
                                intext = true;
                            }
                        

                        //if (runBiblio)
                        
                            String textb = field.Code.Text;
                            String[] biblioSplit = new String[1];
                            biblioSplit[0] = "REFLIST";
                            if (textb.Split(biblioSplit, StringSplitOptions.None).Length > 1)
                            {
                                biblio = true;
                            }
                        

                        if (intext && biblio)
                        {
                            break;
                        }
                    }
                }
            }

        }

        public bool runInUse()
        {
            if (!runSub)
            {
                return false;
            }
            return true;
        }

        public bool runInText()
        {
            if (!intext)
            {
                return false;
            }
            return true;
            
        }

 
        public bool runBiblio()
        {
            if (!biblio)
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
                    Word.Fields f = range.Fields;
                    foreach (Word.Field field in f)
                    {
                        String text = field.Code.Text;
                        String[] intextSplit = new String[1];
                        intextSplit[0] = "EndNote";
                        String[] result = text.Split(intextSplit, StringSplitOptions.None);
                        if (result.Length > 1)
                        {
                            return false;
                        }
                    }
                }
            }
            return true;
        }


    }
}
