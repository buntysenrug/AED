using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class PageMargin:Styles
    {
        private double top;
        private double bottom;
        private double right;
        private double left;

        public PageMargin(Word.Document doc, Word.Application app,double top_margin,double bottom_margin,
            double left_margin,double right_margin)
            : base(doc, app)
        {
            this.top = top_margin;
            this.bottom = bottom_margin;
            this.left = left_margin;
            this.right = right_margin;
        }

        public bool runTop()
        {
            if (top <= 1.5 || top >= 3.2)
            {
                return false;
            }
            return true;
        }

        public bool runBottom()
        {
            if (bottom <= 1.5 || bottom >= 3.2)
            {
                return false;
            }
            return true;
        }

        public bool runLeft()
        {
            if (left <= 1.5 || left >= 3.5)
            {
                return false;
            }
            return true;
        }

        public bool runRight()
        {
            if (right <= 1.5 || right >= 3.2)
            {
                return false;
            }
            return true;
        }


    }
}
