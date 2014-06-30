using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class PsychologyPageBreak:Styles
    {
        private int countNPBAny;
        private int countNPBIntroRef;

        public PsychologyPageBreak(Word.Document doc, Word.Application app,int countPB,int countNPBRef)
            : base(doc, app)
        {
            this.countNPBAny = countPB;
            this.countNPBIntroRef=countNPBRef;
        }

        public bool runCountNPB()
        {
            if (countNPBAny == 0)
            {
                return false;
            }
            return true;
        }

        public bool runNPBRef()
        {
            if (countNPBIntroRef == 0)
            {
                return false;
            }
            return true;
        }


    }
}
