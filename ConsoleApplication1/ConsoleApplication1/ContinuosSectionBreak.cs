using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class ContinuosSectionBreak:Styles
    {
        private bool continousTwice;
        private bool continiousInMiddle;

        public ContinuosSectionBreak(Word.Document doc, Word.Application app, bool continoustwice, bool continuosinmiddle)
            : base(doc,app)
        {
            this.continousTwice = continoustwice;
            this.continiousInMiddle = continuosinmiddle;
        }

        public bool runTwoNext()
        {
            if (continousTwice)
            {
                return false;
            }
            return true;
        }

        public bool runInMiddle()
        {
            if (continiousInMiddle)
            {
                return false;
            }
            return true;
        }


    }
}
