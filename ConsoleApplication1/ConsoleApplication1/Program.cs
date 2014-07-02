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
        {/*
            Word.Application w = new Word.Application();
            Word.Document doc = null;
            doc = w.Documents.Open(@"C:\Users\b1036970\Desktop\ABowen_FINALDISSERTATION1.docx");
            w.Visible = false;
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                float f = p.SpaceAfter;
                Console.WriteLine(p.PageBreakBefore);
            
                
            }

            
            w.Quit();
           */
            String filename="C:\\Users\\b1036970\\Desktop\\ABowen_FINALDISSERTATION1.docx";
            String f = "X:\\LTMS\\TEACHING\\AEDI Software Project\\Testing folders\\Test 4 - MBBS A3\\14627_Upload File_jessica-louise-lugsdin-100495940-14627.docx";
            //String test = "S:\\document.docx";
            Word.Application w = new Word.Application();
            Word.Document doc = null;
            doc = w.Documents.Open(f);
            MBBSA3 mbbs = new MBBSA3(doc, w);
            Dictionary<string,bool> dict=mbbs.initialiseAll();
            foreach (var v in dict)
                Console.WriteLine(v.Key + "   " + v.Value);
            Styles.quit(w, doc);
            Console.WriteLine("End of prog................");
            Console.ReadKey();
            
        }
    }
}
