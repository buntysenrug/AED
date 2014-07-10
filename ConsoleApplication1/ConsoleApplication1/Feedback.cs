using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConsoleApplication1
{
    class Feedback
    {
        private String testName;
        private String feedback_heading;
        private int siteMap;
        private String feedback;

        public Feedback(String tstName, String feedbck, int mapurl, String feedbackHead)
        {
            this.testName = tstName;
            this.feedback = feedbck;
            this.siteMap = mapurl;
            this.feedback_heading = feedbackHead;
        }

        public string getFeedback()
        {
            return this.feedback;
        }
        public string getTestName()
        {
            return this.testName;
        }
        public string getFeedbackHeading()
        {
            return this.feedback_heading;
        }
        public int getMapUrl()
        {
            return this.siteMap;
        }
    }
}
