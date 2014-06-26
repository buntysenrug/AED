using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class SpacingTest:Styles
    {
        private int countPagebreaks;
        private int countNPBAny;
        private int shiftEnters;
        private int tabConsec;
        private int tabStart;
        private int spaceMiddle;
        private int spaceStart;
        private int countDoubleCarriage;
        private int countSingleCarriage;

        public SpacingTest(Word.Document doc,int noOfPageBreaks,int noOfNPB, int noOfShiftEnters, int spaceMiddle,int spaceStart, int noTabconsec,
                            int noOfTabstart, int noOfDoubleCarriage, int noOfSingleCarriage, Word.Application app)
            : base(doc,app)
        {
            this.countPagebreaks = noOfPageBreaks;
            this.countNPBAny = noOfNPB;
            this.shiftEnters = noOfShiftEnters;
            this.tabConsec = noTabconsec;
            this.tabStart = noOfTabstart;
            this.spaceMiddle = spaceMiddle;
            this.spaceStart = spaceStart;
            this.countDoubleCarriage = noOfDoubleCarriage;
            this.countSingleCarriage = noOfSingleCarriage;

        }

        public bool runCarriage()
        {
            if (this.countDoubleCarriage > 0)
            {
                return false;
            }
            return true;
        }

        public bool runCarriageSingle()
        {
            if (countSingleCarriage >= 4)
            {
                return false;
            }
            return true;
        }

        public bool runBreakingMiddle()
        {
            if (spaceMiddle >= 2)
            {
                return false;
            }
            return true;
        }

        public bool runBreakingStart()
        {
            if (spaceStart >= 3)
            {
                return false;
            }
            return true;
        }

        public bool runShiftEnters()
        {
            if (shiftEnters >= 2)
            {
                return false;
            }
            return true;
        }

        public bool runPageBreaks()
        {
            float getPercentage = (countPagebreaks / (countNPBAny + countPagebreaks)) * 100;
            if (getPercentage > 20)
            {
                return false;
            }
            return true;
        }

        public bool runNextBreakAny()
        {
            if (countNPBAny > 0)
            {
                return false;
            }
            return true;
        }

        public bool runTabsStart()
        {
            if (tabStart >= 3)
            {
                return false;
            }
            return true;
        }

        public bool runTabsConsec()
        {
            if (tabConsec >= 2)
            {
                return false;
            }
            return true;
        }
    }
}
