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

        public decimal getHeadingOneMarks()
        {
            int headingtest = getNumberofTest("headingone");
            int headintTrue = getNumberofTrueTest("headingone");
            decimal value = (decimal)headintTrue / headingtest;
            return decimal.Round(value, 2);

        }

        public decimal getHeadingTwoMarks()
        {
            int headingtest = getNumberofTest("headingtwo");
            int headintTrue = getNumberofTrueTest("headingtwo");
            decimal value = (decimal)headintTrue / headingtest;
            return decimal.Round(value, 2);
        }

        public decimal getHeadingThreeMarks()
        {
            int headingtest = getNumberofTest("headingthree");
            int headintTrue = getNumberofTrueTest("headingthree");
            decimal value = (decimal)headintTrue / headingtest;
            return decimal.Round(value, 2);
        }

        public decimal h1TitleNotTwice()
        {
            if (this.dictionary["titleStyleTests_runTitleNotTwice"])
            {
                return 1;
            }
            return 0;
        }

        public decimal getHeadingOrderMarks()
        {
            if (this.dictionary["headingOrderTest_"])
            {
                return 1;
            }
            return 0;
        }

        public decimal getNormalStyleMarks()
        {
            if (dictionary["normalStyleTest_runKeep"])
            {
                int normalstyletest = getNumberofTest("normalstyletest");
                int normalstyleTrue = getNumberofTrueTest("normalstyletest");
                int failTest = normalstyletest - normalstyleTrue;
                if (failTest >= 4)
                {
                    return 0;
                }
                double ratio = 1-(failTest * 0.25);
                
                //decimal value = (decimal)normalstyleTrue / normalstyletest;
                //Console.WriteLine("the value of normal style is " + ratio);
                return (decimal)ratio * 2;
            }
            return 0;
        }

        
    }
}
