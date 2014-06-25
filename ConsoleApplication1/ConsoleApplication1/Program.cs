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
            //String filename="C:\\Users\\b1036970\\Desktop\\ABowen_FINALDISSERTATION1.docx";
            String f = "X:\\LTMS\\TEACHING\\AEDI Software Project\\Testing folders\\Test 4 - MBBS A3\\14732_Upload File_simon-charles-lewis-davies-110055453-14732.docx";
            String test = "S:\\document.docx";
            Word.Application w = new Word.Application();
            Word.Document doc = null;
            doc = w.Documents.Open(test);
            Styles s = new Styles(doc);
            Heading2 h2 = new Heading2(doc);
            Heading1 h1 = new Heading1(doc);
            NormalStyle n = new NormalStyle(doc);
            CaptionStyle c = new CaptionStyle(doc);
            QuoteStyle q = new QuoteStyle(doc);
            ListParagraph ln = new ListParagraph(doc);
            NoSpacingStyle ns = new NoSpacingStyle(doc);
            TitleStyle tts = new TitleStyle(doc);
            ParagraphTest pt = new ParagraphTest(doc);
            SubtitleStyle sub = new SubtitleStyle(doc);

            //Console.WriteLine(n.runBase());
           // Console.WriteLine(h1.runBase());
            Console.WriteLine("The h2 in use is "+h2.runInUse());
            Console.WriteLine(h1.runInUse());
            Console.WriteLine("h1 space after "+h1.runSpaceA());
            Console.WriteLine(c.runInUse());
            Console.WriteLine("Quote Style "+q.runInUse());
            Console.WriteLine(ln.listParaBulletedUsed());
            Console.WriteLine("No Spacing Style test "+ns.noSpacingStyleUsedTest());
            Console.WriteLine("Title style used is "+tts.runTitleUsed());
            pt.iterateOverPara();
            bool subtitle = sub.subTitileStyleUsedTest(pt.getCaptionQuotes());
            Console.WriteLine("Subtitle style is "+sub.subTitileStyleUsedTest(pt.getCaptionQuotes()));
           
            //Styles st = new Styles(doc);
            //st.printStyles();
            
            
            
            //Console.WriteLine(pt.getCaptionQuotes());
            Styles.quit(w, doc);
            Console.WriteLine("End of prog................");
            Console.ReadKey();
            
        }
    }
}
