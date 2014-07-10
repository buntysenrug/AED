using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word=Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            //String filename="C:\\Users\\b1036970\\Desktop\\ABowen_FINALDISSERTATION1.docx";
           //String f = "X:\\LTMS\\TEACHING\\AEDI Software Project\\Testing folders\\Test 4 - MBBS A3\\14627_Upload File_jessica-louise-lugsdin-100495940-14627.docx";


            //var files = System.IO.Directory.GetFiles(folder, "*.*", SearchOption.TopDirectoryOnly)
            //.Where(s => s.EndsWith(".docx") || s.EndsWith(".doc"));
            var filepaths = System.IO.Directory.GetFiles("S:\\Testdocuments", "*.*", System.IO.SearchOption.TopDirectoryOnly).
                Where(s => s.EndsWith(".docx"));
            
            
            
          
            foreach (String s in filepaths)
            {
                bool hiddenFile = System.Text.RegularExpressions.Regex.IsMatch(s, "\\$");
                if (!hiddenFile) 
                    processFile(s);
            }
           
            Console.WriteLine("End of prog................");
            Console.ReadKey();
            
        }

        private static void processFile(string p)
        {
            Word.Application w = new Word.Application();
            Word.Document doc = w.Documents.Open(p);
            w.Visible = false;
            MBBSA3 mbbs = new MBBSA3(doc, w);
            Dictionary<string,bool> dict=mbbs.initialiseAll();

            foreach (var v in dict)
                Console.WriteLine(v.Key + "    " + v.Value);
           //PSY1001 psy = new PSY1001(doc, w);
            //psy.initialiseAll();
            DocumentFeedBack docfeed = new DocumentFeedBack(dict);
            docfeed.printFeedback();
            Marking mark = new Marking(dict);
            Console.WriteLine("the total marks for heading tests is :- " + mbbs.getTotalMarks(dict));
           Console.WriteLine("finished file  " + doc.Name);
           //w.Quit();
           Styles.quit(w, doc);
        }
    }
}
