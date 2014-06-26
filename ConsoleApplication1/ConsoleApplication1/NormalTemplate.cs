using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class NormalTemplate:Styles
    {
        public NormalTemplate(Word.Document doc,Word.Application app)
            : base(doc,app)
        {

        }

        public bool thesisNormalTempTest()
        {
            Word.Template template = doc.get_AttachedTemplate();
            String tempName = template.Name;
            if (tempName.Equals("Normal.dotm") || tempName.Contains("Normal"))
            {
                return false;
            }
            return true;
        }

    }
}
