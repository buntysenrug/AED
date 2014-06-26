using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word=Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class PsychologyTemplate:Styles
    {
        public PsychologyTemplate(Word.Document doc, Word.Application app)
            : base(doc,app)
        {

        }

        public bool psychologyTempTest()
        {
            Word.Template template = doc.get_AttachedTemplate();
            String tempName = template.Name;
            if (tempName.Equals("psychology.dotx") || tempName.Contains("psychology"))
            {
                return true;
            }
            return false;
        }


    }
}
