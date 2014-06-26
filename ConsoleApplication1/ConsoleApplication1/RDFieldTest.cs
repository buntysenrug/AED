using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class RDFieldTest:Styles
    {
        public RDFieldTest(Word.Document doc, Word.Application app)
            : base(doc, app)
        {

        }

        public bool rdFieldsTest()
        {
            Word.Fields f = doc.Fields;
            bool usedRD = false;
            foreach (Word.Field field in f)
            {
                Word.WdFieldType theType = field.Type;

                if (theType == Word.WdFieldType.wdFieldRefDoc)
                {
                    usedRD = true;
                    break;
                }
            }
            return usedRD;
        }
    }
}
