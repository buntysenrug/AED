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

            HashSet<Word.Style> hash = s.getStyles();
            foreach (Word.Style t in hash)
            {
                if(t.NameLocal.Equals("Normal")){
                    Console.WriteLine(t.ParagraphFormat.WidowControl);
                    break;
                }
            }
            Console.ReadKey();
        }
    }
}
