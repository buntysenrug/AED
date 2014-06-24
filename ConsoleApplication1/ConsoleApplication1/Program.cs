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
            String f = "X:\\LTMS\\TEACHING\\AEDI Software Project\\Testing folders\\Test 4 - MBBS A3\\14631_Upload File_alexander-charles-joseph-frost-younger-110080075-14631.docx";
            String test = "S:\\document1.docx";
            Word.Application w = new Word.Application();
            Word.Document doc = null;
            doc = w.Documents.Open(f);
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

            Console.WriteLine(n.runBase());
           // Console.WriteLine(h1.runBase());
            Console.WriteLine("The h2 in use is "+h2.runInUse());
            Console.WriteLine(h1.runInUse());
            Console.WriteLine("h1 space after "+h1.runSpaceA());
            Console.WriteLine(c.runInUse());
            Console.WriteLine("Quote Style "+q.runInUse());
            Console.WriteLine(ln.listParaBulletedUsed());
            Console.WriteLine("No Spacing Style test "+ns.noSpacingStyleUsedTest());
            Console.WriteLine("Title style used is "+tts.runTitleUsed());
            //Styles st = new Styles(doc);
            //st.printStyles();
            pt.iterateOverPara();
            Styles.quit(w, doc);
            Console.ReadKey();
            
        }
    }
}
