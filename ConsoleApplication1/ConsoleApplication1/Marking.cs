using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConsoleApplication1
{
    class Marking
    {
        private Dictionary<String, bool> dictionary;

        public Marking(Dictionary<string, bool> dict)
        {
            this.dictionary = dict;
        }

        public int getNumberofTest(String test)
        {
            int number = 0; ;
            foreach (var v in this.dictionary)
            {

                if (v.Key.IndexOf(test, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    number = number + 1;
                }
            }
            return number;
        }

        public int getNumberofTrueTest(String test)
        {
            int number = 0; ;
            foreach (var v in this.dictionary)
            {

                if (v.Key.IndexOf(test, StringComparison.OrdinalIgnoreCase) == 0 && v.Value == true)
                {
                    number = number + 1;
                }
            }
            return number;

        }

        public decimal getDecimalValueResult(String test)
        {
            int totalTest = getNumberofTest(test);
            int noOftrue = getNumberofTrueTest(test);
            decimal value = (decimal)noOftrue / totalTest;
            return decimal.Round(value, 2);
        }

        public decimal getHeadingMarks()
        {
            int headingtest = getNumberofTest("heading");
            int headintTrue = getNumberofTrueTest("heading");
            decimal value = (decimal)headintTrue / headingtest;
            return decimal.Round(value, 2);

        }
    }
}
