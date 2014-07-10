using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConsoleApplication1
{
    class DocumentFeedBack
    {
        private Dictionary<String, bool> dictionary;
        private Dictionary_Feedback feedback_dict;

        public DocumentFeedBack(Dictionary<string, bool> dict)
        {
            this.dictionary = dict;
            feedback_dict = new Dictionary_Feedback();
        }

        public void printFeedback()
        {
            Dictionary<string, Feedback> feed = feedback_dict.getDict();
            foreach (var v in this.dictionary)
            {
                if (v.Value == false)
                {
                    Feedback f = feed[v.Key];
                    Console.WriteLine(f.getFeedback());
                }
            }
        }
    }
}
