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

        public Dictionary<string,bool> initialiseAll()
        {
            ParagraphTest paratest = new ParagraphTest(doc, app);
            paratest.runDependencies();
            Heading1 head1test = new Heading1(doc, app);
            TitleStyle titletest = new TitleStyle(doc, app);
            Heading2 head2test = new Heading2(doc, app);
            Heading3 head3test = new Heading3(doc, app);
            NormalStyle normal = new NormalStyle(doc, app);
            SubtitleStyle subtitle = new SubtitleStyle(doc, app);
            NoSpacingStyle nospacing = new NoSpacingStyle(doc, app);
           
            CharacterStyle character = new CharacterStyle(doc, app);
            NormalWebStyle normalweb = new NormalWebStyle(doc, app);
            SpacingTest spacetest = new SpacingTest(doc, paratest.getPageBreaks(),
                Convert.ToInt16(paratest.getBPBAny()), paratest.getNoOfShiftEnters(),
                paratest.getSpaceMiddle(), paratest.getSpaceStart(),
                paratest.getTabConsec(), paratest.getTaStart(), 
                paratest.getDoubleCarriage(), paratest.getSingleCarriage(),app);
            QuoteStyle quotetest = new QuoteStyle(doc, app);
            HeaderStyle headerstyle = new HeaderStyle(doc, app);
            FooterStyle footerstyle=new FooterStyle(doc,app);
            StudentNumber studentNo = new StudentNumber(doc, app);
            Images img = new Images(doc, app, paratest.getNumberOfImages(), paratest.getNumberOfFigureCaps());
            Table table = new Table(doc, app, paratest.getLeft(), paratest.getRight(), paratest.getNumberOfTableCaps());
            CaptionRefTest capref = new CaptionRefTest(doc, app, paratest.getNumberOfFigureCaps(), paratest.getNumberOfTableCaps(),
                paratest.getNumberOfImages(), paratest.getNoExplainCaption(), paratest.getAllCaption(), paratest.getCaptionObjects());
            CaptionStyle capstyle = new CaptionStyle(doc, app);
            CrossRefTest crossref = new CrossRefTest(doc, app, paratest.getNumberOfTableCaps(), paratest.getNumberOfFigureCaps(),
                paratest.getNumberOfImages(), paratest.getNumberOfTables(), paratest.getReflink(), paratest.getNumberOfVrossRef());

            TableOfFiguresTest toftest = new TableOfFiguresTest(doc, app);
            //sb.AppendLine(doc.FullName);
            String docname = "S:\\\\TestDoc\\\\"+doc.Name;
            String dd = docname.Substring(0, docname.Length - 4);
            dd = dd + "txt";
            System.IO.StreamWriter file = new System.IO.StreamWriter(dd);
            file.Flush();
            StringBuilder sb = new StringBuilder();
            Styles s = new Styles(doc, app);
            
            //Calling all the methods.
            //header1 methods
           /*sb.AppendLine("Heading 1 run in use is :- " + head1test.runInUse());
            sb.AppendLine("Heading 1 run base is :- " + head1test.runBase());
            sb.AppendLine("Heading 1 run Outline is :- " + head1test.runOutline());
            sb.AppendLine("Heading 1 run keep is :- " + head1test.runKeep());
            sb.AppendLine("Heading 1 run Numbered is :- " + head1test.runNumbered());
            sb.AppendLine("Heading 1 run Bulleted is :- " + head1test.runBulleted());
            sb.AppendLine("Heading 1 run Total space is is :- " + head1test.runTotalSpace());
            //Title test methods
            sb.AppendLine("***********************Title test*************************");
            sb.AppendLine("Title test runtitleNotTwice is :- " + titletest.runTitleNotTwice(paratest.getStylesInDoc(),paratest.getTitleCount()));
            //heading 2 test
            sb.AppendLine("*********************Heading 2 Test***********");
            sb.AppendLine("Heading 2 run in use is :- " + head2test.runInUse());
            sb.AppendLine("Heading 2 run base is :- " + head2test.runBase());
            sb.AppendLine("Heading 2 run Outline is :- " + head2test.runOutline());
            sb.AppendLine("Heading 2 run keep is :- " + head2test.runKeep());
            sb.AppendLine("Heading 2 run Numbered is :- " + head2test.runNumbered());
            sb.AppendLine("Heading 2 run Bulleted is :- " + head2test.runBulleted());
            sb.AppendLine("Heading 2 run Total space is is :- " + head2test.runTotalSpace());
            //heading 3 tests
            sb.AppendLine("*********************Heading 3 Test***********");
            sb.AppendLine("Heading 3 run in use is :- " + head3test.runInUse());
            sb.AppendLine("Heading 3 run base is :- " + head3test.runBase());
            sb.AppendLine("Heading 3 run Outline is :- " + head3test.runOutline());
            sb.AppendLine("Heading 3 run keep is :- " + head3test.runKeep());
            sb.AppendLine("Heading 3 run Numbered is :- " + head3test.runNumbered());
            sb.AppendLine("Heading 3 run Bulleted is :- " + head3test.runBulleted());
            sb.AppendLine("Heading 3 run Total space is is :- " + head3test.runTotalSpace());
            //heading order
            sb.AppendLine("*********************Heading order Test***********");
            sb.AppendLine("Heading order test is :- " + paratest.headingOrderTest());
            //Normal style test
            sb.AppendLine("*********************Normal Style Test***********");
            sb.AppendLine("Normal run in use is :- " + normal.runInUse());
            sb.AppendLine("Normal run base is :- " + normal.runBase());
            sb.AppendLine("Normal run Outline is :- " + normal.runOutline());
            sb.AppendLine("Normal run keep is :- " + normal.runKeep());
            sb.AppendLine("Normal run Font style is :- " + normal.runFontStyle());
            sb.AppendLine("Normal run Font size is :- " + normal.runFontSize());
            sb.AppendLine("Normal run Font effets is :- " + normal.runFontEffects());
            sb.AppendLine("Normal run Total space is is :- " + normal.runTotalSpace());
            //paragraph style test
            sb.AppendLine("*********************Paragraph Style Test***********");
            sb.AppendLine("Paragraph style test is:-  " + paratest.paragraphStyleTest(0));
            //subtitle test
            sb.AppendLine("*********************Subtitle Style Test***********");
            sb.AppendLine("Subtitle Style used is :- " + subtitle.subTitileStyleUsedTest(paratest.getSubtitleQuotes()));
            sb.AppendLine("*********************Character Style Test***********");
            sb.AppendLine("Character Style test is :- " + character.characterStyleTest(paratest.getCharacterQuotes()));
            sb.AppendLine("*********************Normal web  Style Test***********");
            sb.AppendLine("Normal Web style test is " + normalweb.normalWebStyleUsedTest(paratest.getNormalwebQuotes()));
            sb.AppendLine("*********************Spacing Test***********");
            //sb.AppendLine("Spacing test runCarriage is:- ");
            sb.AppendLine("Spacing test runcarriage is:- " + spacetest.runCarriage());
            sb.AppendLine("Spacing test runcarriage single is:- " + spacetest.runCarriageSingle());
            sb.AppendLine("Spacing test runBreakMiddle  is:- " + spacetest.runBreakingMiddle());
            sb.AppendLine("Spacing test runBreakStart is:- " + spacetest.runBreakingStart());
            sb.AppendLine("Spacing test runtabstart is:- " + spacetest.runTabsStart());
            sb.AppendLine("Spacing test runTabconsec is:- " + spacetest.runTabsConsec());
            sb.AppendLine("Spacing test runShiftEnters is:- " + spacetest.runShiftEnters());
            sb.AppendLine("*********************Quote Test***********");
            sb.AppendLine("Quote test runbase is :- " + quotetest.runBase());
            sb.AppendLine("Quote test runFontstyle is :- " + quotetest.runFontStyle());
            sb.AppendLine("Quote test runSpaceA is :- " + quotetest.runSpaceA());
            sb.AppendLine("Quote test runIndent is :- " + quotetest.runIndent());
            sb.AppendLine("*********Header Style Used**************");
            sb.AppendLine("Header Style used test is :- " + headerstyle.headerStyleUsedTest());
            sb.AppendLine("*********Footer Style Used**************");
            sb.AppendLine("Footer Style used test is :- " + footerstyle.footerStyleUsedTest());
            sb.AppendLine("*********Student number Test**************");
            sb.AppendLine("Student Number Test is  :- " + studentNo.studentNumberTest());
            sb.AppendLine("*********Style in Use**************");
            sb.AppendLine("Style in use test is :- " + s.stylesInUseTest());
            Console.WriteLine(s.stylesInUseTest());
            file.Write(sb.ToString());
            file.Close();*/
            Dictionary<string, bool> dictionary = new Dictionary<string, bool>();
            dictionary.Add("headingOneStyleTest_runInUse", head1test.runInUse());
            dictionary.Add("headingOneStyleTest_runBase", head1test.runBase());
            dictionary.Add("headingOneStyleTest_runOutline", head1test.runOutline());
            dictionary.Add("headingOneStyleTest_runKeep", head1test.runKeep());
            dictionary.Add("headingOneStyleTest_runNumbered", head1test.runNumbered());
            dictionary.Add("headingOneStyleTest_runBulleted", head1test.runBulleted());
            dictionary.Add("headingOneStyleTest_runTotalSpace", head1test.runTotalSpace());
            //adding title test
            dictionary.Add("titleStyleTests_runTitleNotTwice" ,titletest.runTitleNotTwice(paratest.getStylesInDoc(), paratest.getTitleCount()));
            //heading 2 test
            dictionary.Add("headingTwoStyleTest_runInUse", head2test.runInUse());
            dictionary.Add("headingTwoStyleTest_runBase", head2test.runBase());
            dictionary.Add("headingTwoStyleTest_runOutline", head2test.runOutline());
            dictionary.Add("headingTwoStyleTest_runKeep", head2test.runKeep());
            dictionary.Add("headingTwoStyleTest_runNumbered", head2test.runNumbered());
            dictionary.Add("headingTwoStyleTest_runBulleted", head2test.runBulleted());
            dictionary.Add("headingTwoStyleTest_runTotalSpace", head2test.runTotalSpace());
            //heading 3 tests
            dictionary.Add("headingThreeStyleTest_runInUse", head3test.runInUse());
            dictionary.Add("headingThreeStyleTest_runBase", head3test.runBase());
            dictionary.Add("headingThreeStyleTest_runOutline", head3test.runOutline());
            dictionary.Add("headingThreeStyleTest_runKeep", head3test.runKeep());
            dictionary.Add("headingThreeStyleTest_runNumbered", head3test.runNumbered());
            dictionary.Add("headingThreeStyleTest_runBulleted", head3test.runBulleted());
            dictionary.Add("headingThreeStyleTest_runTotalSpace", head3test.runTotalSpace());
            //heading order test
            dictionary.Add("headingOrderTest", paratest.headingOrderTest());
            //normal style test
            dictionary.Add("normalStyleTest_runInUse", normal.runInUse());
            dictionary.Add("normalStyleTest_runBase", normal.runBase());
            dictionary.Add("normalStyleTest_runOutline", normal.runOutline());
            dictionary.Add("normalStyleTest_runKeep", normal.runKeep());
            dictionary.Add("normalStyleTest_runFontEffects", normal.runFontEffects());
            dictionary.Add("normalStyleTest_runFontSize", normal.runFontSize());
            dictionary.Add("normalStyleTest_runFontStyle", normal.runFontStyle());
            dictionary.Add("normalStyleTest_runTotalSpace", head3test.runTotalSpace());
            //paragraph test
            dictionary.Add("paragraphStyleTest", paratest.paragraphStyleTest(3));
            //subtitle test
            dictionary.Add("subtitleStyleUsedTest", subtitle.subTitileStyleUsedTest(paratest.getSubtitleQuotes()));
            //character style quotes
            dictionary.Add("characterStyleTest",character.characterStyleTest(paratest.getCharacterQuotes()));
            //normal web style
            dictionary.Add("normalWebStyleUsedTest",normalweb.normalWebStyleUsedTest());
            //spacing tests
            dictionary.Add("spacingTests_runCarriage",spacetest.runCarriage());
            dictionary.Add("spacingTests_runCarriageSingle",spacetest.runCarriageSingle());
            dictionary.Add("spacingTests_runBreakingMiddle",spacetest.runBreakingMiddle());
            dictionary.Add("spacingTests_runBreakingStart", spacetest.runBreakingStart());
            dictionary.Add("spacingTests_runTabStart", spacetest.runTabsStart());
            dictionary.Add("spacingTests_runTabConsec", spacetest.runTabsConsec());
            dictionary.Add("spacingTests_runShiftEnters",spacetest.runShiftEnters());
            //quote tests
            dictionary.Add("quoteStyleTest_runBase", quotetest.runBase());
            dictionary.Add("quoteStyleTest_runFontStyle", quotetest.runFontStyle());
            dictionary.Add("quoteStyleTest_runSpaceA", quotetest.runSpaceA());
            dictionary.Add("quoteStyleTest_runIndent", quotetest.runIndent());
            //header style test
            dictionary.Add("headerStyleUsedTest", headerstyle.headerStyleUsedTest());
            //footer style test
            dictionary.Add("footerStyleUsedTest", footerstyle.footerStyleUsedTest());
            //student number test
            dictionary.Add("studentNumberTest", studentNo.studentNumberTest());
            //styles in use test
            dictionary.Add("stylesInUseTest", paratest.stylesInUseTest());
            
            return dictionary;
        }

        public decimal getTotalMarks(Dictionary<string,bool> results)
        {
            Marking mark = new Marking(results);

            return mark.getHeadingMarks();
        }

       
    }
}
