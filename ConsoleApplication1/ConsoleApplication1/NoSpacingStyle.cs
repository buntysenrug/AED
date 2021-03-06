﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class NoSpacingStyle:Styles
    {
        private Word.Style nospace;
        public NoSpacingStyle(Word.Document doc,Word.Application app)
            : base(doc,app)
        {
            foreach (Word.Style s in set)
            {
                if (s.NameLocal.Contains("No Spacing") || s.NameLocal.Equals("No Spacing"))
                {
                    this.nospace = s;
                    break;
                }
            }

        }

        public bool noSpacingStyleUsedTest()
        {
            try
            {
                return !this.nospace.InUse;
            }
            catch (Exception ex)
            {
                return true;
            }
        }
    }
}
