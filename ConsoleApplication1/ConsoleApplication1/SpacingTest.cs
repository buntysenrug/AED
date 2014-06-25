using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class SpacingTest:Styles
    {
        public SpacingTest(Word.Document doc)
            : base(doc)
        {

        }

        public bool runCarriage(int doubleCarriage)
        {
            if (doubleCarriage > 0)
            {
                return false;
            }
            return true;
        }

        public bool runCarriageSingle(int singleCarriage)
        {
            return true;
        }
    }
}
