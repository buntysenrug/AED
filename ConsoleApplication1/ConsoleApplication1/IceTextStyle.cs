using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class IceTextStyle:Styles
    {
        private Word.Style icestyle;
        private float spaceBeforeLower;
        private float spaceBeforeUpper;
        private float spaceAfterLower;
        private float spaceAfterUpper;

        public IceTextStyle(Word.Document doc, Word.Application app)
            : base(doc,app)
        {
            
            foreach (Word.Style s in set)
            {
                if (s.NameLocal.Equals("Ice Text") || s.NameLocal.Contains("Ice Text"))
                {
                    icestyle = s;
                    break;
                }
            }
            this.spaceBeforeLower = 0f;
            this.spaceBeforeUpper = 12f;
            this.spaceAfterLower = 6f;
            this.spaceAfterUpper = 18f;

        }

        public bool runInUse()
        {
            try
            {
                return icestyle.InUse;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        public bool runBase()
        {
            try
            {
                Word.Style base_style = getBaseStyle(icestyle.NameLocal);
                if (base_style.NameLocal.Equals("Normal"))
                {
                    return true;
                }
                return false;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        public bool runBorder()
        {
            try
            {
                String descrip = icestyle.Description;
                String[] splitter = new String[1];
                splitter[0] = "Box";
                splitter = descrip.Split(splitter, StringSplitOptions.None);
                if (splitter.Length > 1)
                {
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool runSpaceB()
        {
            try
            {
                return spaceBeforeStyleCheck(icestyle, this.spaceBeforeLower, this.spaceBeforeUpper);
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool runSpaceA()
        {
            try
            {
                return spaceAfterStyleCheck(icestyle, this.spaceAfterLower, this.spaceAfterUpper);
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool QuickA()
        {
            try
            {
                if (!inQuickStyleListCheck(icestyle, true))
                {
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }




    }
}
