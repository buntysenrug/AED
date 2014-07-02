using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word=Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class MBBSA3
    {
        private Word.Document doc;
        private Word.Application app;

        public MBBSA3(Word.Document d, Word.Application a)
        {
            this.doc = d;
            this.app = a;
        }

        public void initialiseAll()
        {
            ParagraphTest paratest = new ParagraphTest(doc, app);
            Heading1 head1test = new Heading1(doc, app);
            TitleStyle titletest = new TitleStyle(doc, app);
            Heading2 head2test = new Heading2(doc, app);
            Heading3 head3test = new Heading3(doc, app);
            NormalStyle normal = new NormalStyle(doc, app);
            SubtitleStyle subtitle = new SubtitleStyle(doc, app);
            NoSpacingStyle nospacing = new NoSpacingStyle(doc, app);
            CharacterStyle character = new CharacterStyle(doc, app);
            NormalWebStyle normalweb = new NormalWebStyle(doc, app);
            //SpacingTest spacetest = new SpacingTest(doc,paratest.n
            QuoteStyle quotetest = new QuoteStyle(doc, app);
        }
    }
}
