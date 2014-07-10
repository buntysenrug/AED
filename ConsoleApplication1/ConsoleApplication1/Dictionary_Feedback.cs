using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace ConsoleApplication1
{
    class Dictionary_Feedback
    {
        private Dictionary<string, Feedback> dictionary;
        public Dictionary_Feedback()
        {
            dictionary = new Dictionary<string, Feedback>();
            XmlDocument doc = new XmlDocument();
            doc.Load("C:\\Users\\b1036970\\Desktop\\feedback.xml");
            XmlNode x = doc.SelectSingleNode("Feedback");
            foreach (XmlNode child in x)
            {
                string testname = child.Attributes["testName"].Value;
                string feedback_head = child.Attributes["feedBackHeading"].Value;
                int siteUrl = Convert.ToInt16(child.Attributes["siteUrl"].Value);
                string feedback = child.Attributes["feedback"].Value;

                dictionary.Add(child.Name + "_" + testname, new Feedback(testname, feedback, siteUrl, feedback_head));
            }

        }

        public void print()
        {
            foreach (var i in this.dictionary)
            {
                Console.WriteLine(i.Key);
                Console.WriteLine(i.Value.getFeedback());
                Console.WriteLine(i.Value.getMapUrl());
            }
        }

        public Dictionary<string, Feedback> getDict()
        {
            return this.dictionary;
        }
    }
}
