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
            Styles s = new Styles(filename);

            Dictionary<Word.Style, Word.Style> dict = s.getBaseStyles();
            foreach (var entry in dict)
                Console.WriteLine("[{0} {1}]", entry.Key.NameLocal, entry.Value.NameLocal); 
            Console.ReadKey();
        }
    }
}
