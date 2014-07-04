using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word=Microsoft.Office.Interop.Word;
using Style=Microsoft.Office.Interop.Word.Style;

namespace ConsoleApplication1
{
    class ParagraphTest : Styles
    {
        private Style quoteStyle;
        private bool headingOrderError;
        private Style quotechar;
        private bool foundQuote;
        private Style emphasis;
        private Style strong;
        private Style hyperStyle;
        private bool gotEmphasis;
        private bool gotStrong;
        private bool matchPlagState;
        private Style h1style;
        private Style tofstyle;
        private Style emphStyle;
        private Style strongStyle;
        private bool gotTofStyle;
        private String endNoteResult;
        private bool gotMargins;
        private bool continueTitleSearch;
        private int lastHeading;
        private bool gotH1;
        private bool captionOK;
        private bool matchTempTip;
        private bool titleUsedThree;
        private bool TOFStyleUsed;
        private bool refStyleUsed;
        private bool usedNumberedList;
        private Style[] styles;
        private int countSingleCarriage;
        private int countPageBreaks;
        private float countNPBAny;
        private int countDoubleCarriage;
        private int carriageCounter;
        private int carriageOccasions;
        private int titleCount;
        private int countNPBIntroRef;
        private bool usedFASEBStyle;
        private bool usedBulletedList;
        private double leftMargin;
        private double rightMargin;
        private double topMargin;
        private double bottomMargin;
        private int refLinkBroken;
        private bool columnsError;
        private bool continousTwice;
        private bool continiousInMiddle;
        private bool alreadyGotStyles;
        private bool usedAPAStyle;
        private bool noMutli;
        private bool multiError;
        private int numOfFigureCaps;
        private int numOfTableCaps;
        private int numberOfTables;
        private int numberOfImages;
        private int numOfCrossRef;

        //List
        private List<String> allParagraphText;
        private List<Word.Style> stylesInDoc;
        private List<String> subtitleStyleQuotes;
        private List<String> noSpacingStyleQuotes;
        private List<String> headerStyleQuotes;
        private List<String> footerStyleQuotes;
        private List<String> footNoteStyleQuotes;
        private List<String> normalWebStyleQuotes;
        private List<String> captionWithObjects;
        private List<String> noExplanCaptionText;
        private List<String> allCaptionText;
        private List<String> paragraphStyles;
        private List<String> captionQuotes;
        private List<string> shiftEnters;
        private List<String> characterStyleQuotes;
        //---
        private List<Word.Paragraph> spaceMiddle;
        private List<Word.Paragraph> spaceStart;
        private List<Word.Paragraph> tabsConsec;
        private List<Word.Paragraph> tabsStart;


        public ParagraphTest(Word.Document doc, Word.Application app)
            : base(doc,app)
        {
            this.gotMargins = false;
            this.continueTitleSearch = true;
            this.lastHeading = 0;
            this.headingOrderError = false;
            this.quoteStyle = getStyle("Quote");
            this.quotechar = quoteStyle.get_LinkStyle();
            this.foundQuote = false;
            this.emphasis = getStyle("Emphasis");
            this.strong = getStyle("Strong");
            this.hyperStyle = getStyle("Hyperlink");
            this.gotEmphasis = false;
            this.gotStrong = false;
            this.gotH1 = false;
            this.matchTempTip = false;
            this.matchPlagState = false;
            this.h1style = getStyle("Heading 1");
            this.tofstyle = getStyle("Table of Figures");
            this.gotTofStyle = false;
            this.continueTitleSearch = true;
            this.titleUsedThree = false;
            this.countSingleCarriage = 0;
            this.countPageBreaks = 0;
            this.countDoubleCarriage = 0;
            this.carriageCounter = 0;
            this.carriageOccasions = 0;
            this.titleCount = 0;
            this.TOFStyleUsed = false;
            this.refStyleUsed = false;
            this.usedFASEBStyle = false;
            this.usedNumberedList = false;
            this.usedBulletedList = false;
            this.refLinkBroken = 0;
            this.columnsError = false;
            this.alreadyGotStyles = false;
            this.usedAPAStyle=false;
            this.noMutli = false;
            this.multiError = false;
            this.countNPBAny = 0;
            this.continiousInMiddle = false;
            this.countNPBIntroRef = 0;
            this.numberOfImages=0;
            this.numberOfTables=0;
            this.numOfFigureCaps=0;
            this.numOfTableCaps=0;
            //Initialise List

            allParagraphText = new List<string>();
            stylesInDoc = new List<Word.Style>();
            subtitleStyleQuotes = new List<string>();
            noSpacingStyleQuotes = new List<string>();
            headerStyleQuotes = new List<string>();
            footerStyleQuotes = new List<string>();
            footNoteStyleQuotes = new List<string>();
            normalWebStyleQuotes = new List<string>();
            captionWithObjects = new List<string>();
            noExplanCaptionText = new List<string>();
            allCaptionText = new List<string>();
            paragraphStyles = new List<string>();
            captionQuotes = new List<string>();
            shiftEnters = new List<string>();
            characterStyleQuotes = new List<string>();

            //--
            spaceMiddle = new List<Word.Paragraph>();
            spaceStart = new List<Word.Paragraph>();
            tabsStart = new List<Word.Paragraph>();
            tabsConsec = new List<Word.Paragraph>();

            
        }

        private void iterateOverPara()
        {
            foreach (Word.Paragraph p in doc.Paragraphs)
            {

                //If the current fields in this range have the reflist then we are at endnotes reference list
                foreach (Word.Field field in p.Range.Fields)
                {
                    if (field.Code.Text.Equals(" ADDIN EN.REFLIST "))
                    {
                        endNoteResult = field.Result.Text;///add to this string so can ignore endnote text formatting in later tests
                    }
                }

                if (!inEndnotes(p))
                {
                    allParagraphText.Add(p.Range.Text);//add this paragraph to all text

                    //searches for guide in psy template
                    //if (marker.getProgramme().Equals("psymark1001"))  ***commented out to stop it being test dependent
                    //{
                    bool matchTempTip = System.Text.RegularExpressions.Regex.IsMatch(p.Range.Text, "Tips for using this Template");
                    if (matchTempTip)
                    {
                        this.matchTempTip = true;
                    } 

                    //searches for plagiarism statement in psy document
                    bool matchPlagState = System.Text.RegularExpressions.Regex.IsMatch(p.Range.Text, "I give permission to have my work submitted to an electronic plagiarism checker");
                    if (matchPlagState)
                    {
                        this.matchPlagState = true;
                    }
                    //}


                    spacingPreChecks(p);//run the spacing checks

                    Style paraStyle = p.get_Style();
                    Style charStyle = paraStyle.get_LinkStyle();
                    Word.Range prange = p.Range;

                    //prange.Find.ClearFormatting();
                    //prange.Find.set_Style(tofstyle);
                    //gotTofStyle = prange.Find.Execute();

                    //if (gotTofStyle)
                    //{
                    //    stylesInDoc.Add(tofstyle);
                    //}

                    if (h1style != null)
                    {
                        prange.Find.ClearFormatting();
                        prange.Find.set_Style(h1style);
                        gotH1 = prange.Find.Execute();
                    }

                    if (gotEmphasis)
                    {
                        stylesInDoc.Add(emphasis);
                    }

                    if (!gotEmphasis)
                    {
                        prange.Find.ClearFormatting();
                        prange.Find.set_Style(emphasis);
                        gotEmphasis = prange.Find.Execute();
                        if (gotEmphasis)
                        {
                            stylesInDoc.Add(emphasis);
                        }
                    }

                    if (!gotStrong)
                    {
                        prange.Find.ClearFormatting();
                        prange.Find.set_Style(strong);
                        gotStrong = prange.Find.Execute();
                        if (gotStrong)
                        {
                            stylesInDoc.Add(strong);
                        }
                    }

                    if (!foundQuote)
                    {
                        prange.Find.ClearFormatting();
                        if (quotechar.NameLocal.Equals("Quote Char"))
                        {
                            prange.Find.set_Style(quotechar);
                            foundQuote = prange.Find.Execute();
                            if (foundQuote)
                            {
                                stylesInDoc.Add(quoteStyle);
                                foundQuote = true;
                            }
                        }
                    }


                    //Determine if headings are in correct order
                    if (!p.Range.Text.Equals("\r"))//ignore blank headings
                    {
                        if (!headingOrderError)//keep checking while no problems
                        {
                            bool matchHeadingType = System.Text.RegularExpressions.Regex.IsMatch(paraStyle.NameLocal, "Heading \\d+$");
                            if (matchHeadingType)
                            {
                                String[] headingSplitter = new String[1];
                                headingSplitter[0] = "Heading";
                                headingSplitter = paraStyle.NameLocal.Split(headingSplitter, StringSplitOptions.None);

                                if (headingSplitter.Length > 1)
                                {
                                    int currentNumber = int.Parse(headingSplitter[1]);
                                    if (!(lastHeading >= currentNumber) && lastHeading != currentNumber - 1 && currentNumber != 1)
                                    {
                                        headingOrderError = true;
                                    }
                                    lastHeading = currentNumber;
                                }
                            }
                            else if (paraStyle.NameLocal.Equals("Title"))
                            {
                                lastHeading = 1;
                            }
                        }
                    }

                    if (paraStyle != null)
                    {
                        checkForMulti(p);

                        //checks for list paragraph
                        if (paraStyle.NameLocal.Equals("List Paragraph"))
                        {
                            Word.ListFormat list = p.Range.ListFormat;
                            bool matchNum = System.Text.RegularExpressions.Regex.IsMatch(list.ListString, "\\d");
                            if (matchNum)
                            {
                                usedNumberedList = true;
                            }
                            else
                            {
                                bool matchBull = System.Text.RegularExpressions.Regex.IsMatch(list.ListString, ".");
                                if (matchBull)
                                {
                                    usedBulletedList = true;
                                }
                            }
                        }

                        if (!gotMargins)
                        {
                            leftMargin = Math.Round((double)app.PointsToCentimeters(p.Range.PageSetup.LeftMargin), 3);
                            rightMargin = Math.Round((double)app.PointsToCentimeters(p.Range.PageSetup.RightMargin), 3);
                            topMargin = Math.Round((double)app.PointsToCentimeters(p.Range.PageSetup.TopMargin), 3);
                            bottomMargin = Math.Round((double)app.PointsToCentimeters(p.Range.PageSetup.BottomMargin), 3);
                            
                            gotMargins = true;
                        }

                        //check some other style elements, thse are used later in other methods for style checking, best used here as iterating through ALL paragraphs 
                        if (paraStyle.NameLocal.Equals("Subtitle"))
                        {
                            if (!p.Range.Text.Equals("\r"))
                            {
                                subtitleStyleQuotes.Add(p.Range.Text);
                            }
                        }
                        else if (paraStyle.NameLocal.Equals("No Spacing"))
                        {
                            if (!(p.Range.Tables.Count > 0) && !p.Range.Text.Equals("\r"))
                            {
                                noSpacingStyleQuotes.Add(p.Range.Text);
                            }
                        }
                        else if (paraStyle.NameLocal.Equals("Header"))
                        {
                            headerStyleQuotes.Add(p.Range.Text);
                        }
                        else if (paraStyle.NameLocal.Equals("Footer"))
                        {
                            footerStyleQuotes.Add(p.Range.Text);
                        }
                        else if (paraStyle.NameLocal.Equals("Footnote Text"))
                        {
                            footNoteStyleQuotes.Add(p.Range.Text);
                        }
                        else if (paraStyle.NameLocal.Equals("Normal (Web)"))
                        {
                            if (!p.Range.Text.Equals("\r"))
                            {
                                normalWebStyleQuotes.Add(p.Range.Text);
                            }
                        }
                        else if (paraStyle.NameLocal.Equals("Table of Figures") || paraStyle.NameLocal.Contains("TOC"))
                        {
                            TOFStyleUsed = true;
                        }
                        else if (paraStyle.NameLocal.Equals("References"))
                        {
                            refStyleUsed = true;
                        }
                        else if (paraStyle.NameLocal.Equals("FASEB J"))
                        {
                            usedFASEBStyle = true;
                        }

                        if (p.Range.Fields.Count > 0)
                        {
                            foreach (Word.Field f in p.Range.Fields)
                            {
                                if (f.Type == Word.WdFieldType.wdFieldSequence)
                                {
                                    captionOK = true;
                                    String quote = p.Range.Text;
                                    Word.Paragraph next = p.Next();
                                    int nextParaShapes = 0;
                                    int nextParaTables = 0;
                                    if (next != null)
                                    {
                                        nextParaShapes = next.Range.ShapeRange.Count;
                                        if (nextParaShapes > 0)
                                        {
                                            Word.ShapeRange srange = next.Range.ShapeRange;
                                            for (int i = 1; i <= nextParaShapes; i++)
                                            {
                                                Word.Shape shape = srange[i];
                                                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                                {
                                                    Word.GroupShapes group = shape.GroupItems;
                                                    for (int j = 1; j <= group.Count; j++)
                                                    {
                                                        Word.Shape innershape = group[i];
                                                        if (innershape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                                                        {
                                                            nextParaShapes = 1;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (nextParaShapes == 0)
                                        {
                                            nextParaShapes = next.Range.InlineShapes.Count;
                                            if (nextParaShapes == 0)
                                            {
                                                if (next.Range.Text.Equals("\r"))
                                                {
                                                    Word.Paragraph next2 = next.Next();
                                                    if (next2 != null)
                                                    {
                                                        nextParaShapes = next2.Range.InlineShapes.Count;
                                                        if (nextParaShapes == 0)
                                                        {
                                                            nextParaShapes = next2.Range.ShapeRange.Count;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        nextParaTables = next.Range.Tables.Count;
                                        if (nextParaTables == 0)
                                        {
                                            if (next.Range.Text.Equals("\r"))
                                            {
                                                Word.Paragraph next2 = next.Next();
                                                if (next2 != null)
                                                {
                                                    nextParaTables = next2.Range.Tables.Count;
                                                }
                                            }
                                        }

                                        int thisIndex = p.Range.Sections.First.Index;
                                        int nextIndex = next.Range.Sections.First.Index;

                                        bool hasShapes = p.Range.InlineShapes.Count > 0;//checks if has an inline image 
                                        bool tableCount = p.Range.Tables.Count > 0;// check for tables
                                        String[] paraMark = quote.Split('\r');
                                        bool isParaMark = paraMark[0].Equals("");
                                        bool isBreak = !(thisIndex == nextIndex);

                                        if (hasShapes || tableCount || isParaMark || isBreak)
                                        {
                                            captionOK = false;
                                            bool matchNext = System.Text.RegularExpressions.Regex.IsMatch(p.Next().Range.Text, "\\w+\\s+\\d+");
                                            if (matchNext)
                                            {
                                                captionWithObjects.Add(p.Next().Range.Text);
                                            }
                                        }
                                    }
                                    if (captionOK)
                                    {
                                        String[] removeCarriage = quote.Split('\r');
                                        // bool isMatch = System.Text.RegularExpressions.Regex.IsMatch(removeCarriage[0], "\\w*\\s*\\d+.{1}\\w+");
                                        bool isMatch = System.Text.RegularExpressions.Regex.IsMatch(removeCarriage[0], "\\d.{1}\\D+|\\d.\\d\\D.{1}\\w+");
                                        // Debug.Write(removeCarriage[0]); 
                                        if (!isMatch)
                                        {
                                            noExplanCaptionText.Add(quote);//seperate list for captions that have no explanation 
                                        }

                                        allCaptionText.Add(p.Range.Text);//add all captions to a list
                                        if (nextParaShapes == 0 && nextParaTables == 0)
                                        {
                                            captionQuotes.Add(p.Range.Text);
                                        }
                                    }
                                }
                            }
                        }

                        bool linkBroken = System.Text.RegularExpressions.Regex.IsMatch(p.Range.Text, "Error! Reference source not found.");
                        if (linkBroken)
                        {
                            refLinkBroken++;
                        }


                        bool styleUsed = false;
                        foreach (Style currentStyles in stylesInDoc)
                        {
                            if (currentStyles.NameLocal.Equals(paraStyle.NameLocal))
                            {
                                styleUsed = true;
                                break;
                            }
                        }

                        if (continueTitleSearch)
                        {
                            if (paraStyle.NameLocal.Equals("Title"))
                            {
                                titleCount++;
                                Word.Paragraph secondPara = p.Next();
                                if (secondPara != null)
                                {
                                    Style secondParaStyle = secondPara.get_Style();
                                    if (secondParaStyle.NameLocal.Equals("Title"))
                                    {
                                        titleUsedThree = true;
                                        Word.Paragraph thirdPara = secondPara.Next();
                                        Style thirdParaStyle = thirdPara.get_Style();
                                        if (thirdParaStyle.NameLocal.Equals("Title"))
                                        {

                                            titleUsedThree = true;
                                            Word.Paragraph fourthPara = thirdPara.Next();
                                            Style fourthParaStyle = fourthPara.get_Style();
                                            if (fourthParaStyle.NameLocal.Equals("Title"))
                                            {
                                                titleUsedThree = true;
                                                continueTitleSearch = false;

                                                Word.Paragraph fifthPara = fourthPara.Next();
                                                Style fifthParaStyle = fifthPara.get_Style();
                                                if (fifthParaStyle != null)
                                                {
                                                    if (fifthParaStyle.NameLocal.Equals("Title"))
                                                    {
                                                        titleUsedThree = false;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (!styleUsed)
                        {
                            stylesInDoc.Add(paraStyle);
                        }

                        String text = p.Range.Text;

                        //Check the character style of this paragraph
                        characterStyleSetup(p);

                        //Since we are iterating through paragraphs might as well check the style of this paragraph
                        checkParagraphStyle(p, paraStyle);

                        //check the columns of each paragraph range
                        if (!columnsError)
                        {
                            Word.TextColumns t = p.Range.PageSetup.TextColumns;
                            if (t.Count != 2)
                            {
                                columnsError = true;
                            }
                        }

                    }
                }
                else
                {
                    carriageCounter = 0;
                }
            }

            foreach (Word.Shape s in doc.Shapes)
            {
                if (s.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                {
                    Word.Range range = s.TextFrame.ContainingRange;
                    Style style = range.get_Style();
                    if (style != null)
                    {
                        bool styleUsed = false;

                        foreach (Style currentStyles in stylesInDoc)
                        {
                            if (currentStyles.NameLocal.Equals(style.NameLocal))
                            {
                                styleUsed = true;
                                break;
                            }
                        }
                        if (!styleUsed)
                        {
                            stylesInDoc.Add(style);
                        }
                    }
                }
            }
        }

        /*
        * Tests if the paragraph is part of the endnotes, may be used to avoid certain tests on this section of text
        */
        public bool inEndnotes(Word.Paragraph para)
        {
            if (endNoteResult != null)
            {
                String[] splitter2 = new String[1];
                splitter2[0] = para.Range.Text;
                splitter2 = endNoteResult.Split(splitter2, StringSplitOptions.None);
                if (splitter2.Length == 2)
                {
                    return true;
                }
            }
            return false;
        }

        /*
         *Retrieves the style object based on param name
         */
        private Style getStyle(String name)
        {

            foreach (Style s in this.doc.Styles)
            {

                if (s.NameLocal.Equals(name))
                {
                    return s;
                }
            }
           // Debug.Write("No Style by name of " + name);
            return null;

        }

        /*
         * Used in spacing test, is ran during iterateoverPara test, saves iterating through doc multiple times
         */
        private void spacingPreChecks(Word.Paragraph para)
        {
            bool endnoteTest = inEndnotes(para);
            int startSpaceCount = 0;
            int startSpaceOcc = 0;
            int spaceCounter = 0;
            int spaceOccasions = 0;

            int NumtabConsec = 0;
            int lastSection = 1;

            if (!endnoteTest)
            {
                Word.Paragraph nextP = para.Next();
                Word.Paragraph prevP = para.Previous();
                if (nextP != null && prevP != null)
                {
                    if (para.Range.Text.Equals("\r") && !nextP.Range.Text.Equals("\r") && !prevP.Range.Text.Equals("\r"))
                    {
                        countSingleCarriage++;
                    }
                }

                bool carOK = checkCarriages(para);

                if (!carOK)
                {
                    countDoubleCarriage++;
                }


                bool gotNextPageBreak = System.Text.RegularExpressions.Regex.IsMatch(para.Range.Text, "\\f$");

                //count next page breaks
                if (gotNextPageBreak)
                {
                    countNPBAny++;
                }

                if (gotNextPageBreak)
                {

                    Word.Paragraph next = recursiveGetPara(para);
                    if (next != null)
                    {
                        Style nextParaStyle = next.get_Style();
                        bool match = System.Text.RegularExpressions.Regex.IsMatch(next.Range.Text, "References");
                        if (nextParaStyle.NameLocal.Equals("Heading 1") && match)
                        {
                            countNPBIntroRef++;
                            countNPBAny--;
                        }
                        else if (next.Range.Fields.Count > 0)
                        {
                            foreach (Word.Field f in next.Range.Fields)
                            {
                                if (f.Type == Word.WdFieldType.wdFieldAddin)
                                {
                                    if (f.Code.Text.Equals(" ADDIN EN.REFLIST "))
                                    {
                                        countNPBIntroRef++;
                                        countNPBAny--;
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }

                //checks the spacing middle and spacing start
                String[] splitPara = para.Range.Text.Split(' ');
                int cutOfEnd = 0;
                for (int i = splitPara.Length - 1; i >= 0; i--)
                {
                    cutOfEnd = i;
                    if (!splitPara[i].Equals("") && !splitPara[i].Equals("\r"))
                    {
                        break;
                    }

                }
                int spaceLimit = 4;
                bool foundWords = false;

                //iterate through the paragraph split by \s
                for (int i = 0; i <= cutOfEnd; i++)
                {
                    if (splitPara[i].Equals(""))//found a space
                    {
                        if (foundWords)//if found words at begging i.e found other characters previously
                        {
                            spaceCounter++;//must be in the middle of para so increase spaceCounter
                        }
                        else
                        {
                            startSpaceCount++;//else we are at beggining so increase this counter
                        }
                    }
                    else//not found a space char
                    {
                        spaceCounter = 0;//set back to zero as not found 4
                        startSpaceCount = 0;//set back to zero not found 4 at start 
                        foundWords = true;//found some words so set to true, not at begging anymore
                    }

                    if (spaceCounter == spaceLimit)//reached counter limit (4)
                    {
                        spaceOccasions++;//counts occurances of 4 spaces in middle
                        spaceMiddle.Add(para);//add this paragraph to list tracking paragraphs with spaces in middle
                    }

                    if (startSpaceCount == spaceLimit)//reached counter limit (4)
                    {
                        startSpaceOcc++;//counts occurances of 4 spaces at start 
                        spaceStart.Add(para);//add this paragraph to list tracking paragraphs with spaces in start
                    }
                }


                //testing for tabs 
                String[] splitTabs = para.Range.Text.Split('\t');//split by tab 

                for (int i = 0; i < splitTabs.Length; i++)
                {
                    if (splitTabs[i].Equals(""))//found tab 
                    {
                        NumtabConsec++;//counting consecutive tab chars                         
                    }
                    else
                    {
                        if (i < 3 && NumtabConsec > 0)//found a tab character at start of paragraph but is and does not have 4 consecutive tabs
                        {
                            tabsStart.Add(para);//add to list where tabs are at start 
                        }
                        NumtabConsec = 0; //reset conectutive tabs 
                    }

                    if (NumtabConsec == 4)
                    {
                        tabsConsec.Add(para);//add to list where tabs are consecutive 
                        NumtabConsec = 0;
                    }
                }

                //testing for shift enters
                String[] splitEnters = para.Range.Text.Split('\v');//search for shift enters
                if (splitEnters.Length > 1)
                {
                    if (!splitEnters[0].Equals(""))
                    {
                        shiftEnters.Add(splitEnters[0]);
                    }
                }

                String text = para.Range.Text;
                String[] find = new String[1];
                find[0] = "\f\r";

                String[] pageBreaks = text.Split(find, StringSplitOptions.None);
                if (pageBreaks.Length >= 2)
                {
                    countPageBreaks++;
                }
                lastSection = para.Range.Sections[1].Index;


                String text2 = para.Range.Text;
                String[] spliter = para.Range.Text.Split('\f');
                if (!continiousInMiddle)
                {
                    if (spliter.Length >= 2 && !para.Range.Text.Equals("\f") && !spliter[spliter.Length - 1].Equals(""))
                    {
                        if (!spliter[1].Equals("\r"))
                        {
                            continiousInMiddle = true;

                        }
                    }
                }
                if (!continousTwice)
                {
                    Word.Paragraph next = para.Next();
                    if (next != null)
                    {
                        String[] splitNext = next.Range.Text.Split('\f');

                        if (splitNext.Length >= 2 && spliter.Length >= 2)
                        {
                            if (spliter[spliter.Length - 1].Equals("") && splitNext[splitNext.Length - 2].Equals(""))
                            {
                                continousTwice = true;
                            }
                        }
                    }
                }
            }
        }


        public Word.Paragraph recursiveGetPara(Word.Paragraph para)
        {
            Word.Paragraph next = para.Next();
            if (next == null)
            {
                return next;
            }
            if (!next.Range.Text.Equals("\r"))
            {
                return next;
            }
            else
            {
                return recursiveGetPara(next);
            }
        }

        /*
         * Checks to ensure certain character styles are not used
         */
        private bool characterStyleSetup(Word.Paragraph para)
        {
            Word.Range r = para.Range;
            String[] theStyles = new String[8];
            theStyles[0] = "Subtle Emphasis";
            theStyles[1] = "Intense Emphasis";
            theStyles[2] = "Intense Quote";
            theStyles[3] = "Subtle Reference";
            theStyles[4] = "Intense Reference";
            theStyles[5] = "Book Title";
            //subtitle and intense quote are both paragraph and character styles, therefore need to get link stlye later, here I am just padding out the array to include an extra two elements
            theStyles[7] = "Intense Quote Char";
            Style name = para.get_Style();

            if (!alreadyGotStyles)
            {
                styles = getStyle(theStyles);
                alreadyGotStyles = true;
                //link style method will get the character style that is related to another style, here getting character style of subtitle and intense quote
                foreach (Style paraStyle in styles)
                {
                    if (paraStyle != null)
                    {
                        if (paraStyle.NameLocal.Equals("Intense Quote"))
                        {
                            Style intenseChar = paraStyle.get_LinkStyle();
                            if (!intenseChar.NameLocal.Equals("Normal"))
                            {
                                styles[7] = intenseChar;
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < styles.Length; i++)
            {
                r.Find.ClearFormatting();
                if (styles[i] != null)
                {
                    r.Find.set_Style(styles[i]);
                    bool styleUsed = r.Find.Execute();
                    if (styleUsed)
                    {
                        characterStyleQuotes.Add(para.Range.Text);
                    }
                }
            }
            return true;

        }

        /*
        * Runs checks on each paragraph checking for multilevel lists
        */
        private void checkForMulti(Word.Paragraph p)
        {
            Word.ListFormat format = p.Range.ListFormat;
            Style paraStyle = p.get_Style();
            String styleName = paraStyle.NameLocal;
            if (styleName.Equals("Heading 1") || styleName.Equals("Heading 2") || styleName.Equals("Heading 3") || styleName.Equals("Heading 4") || styleName.Equals("Heading 5") || styleName.Equals("Heading 6"))
            {
                if (format.ListString.Equals(""))
                {
                    noMutli = true;
                }

            }
            if (styleName.Equals("Heading 1"))
            {

                String listString = format.ListString;
                String[] splitter = new String[1];
                splitter[0] = "Chapter";
                splitter = listString.Split(splitter, StringSplitOptions.None);
                if (!(splitter.Length > 0))
                {
                    multiError = true;
                }
                if (paraStyle.ParagraphFormat.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevel1)
                {
                    multiError = true;
                }
            }
            else if (styleName.Equals("Heading 2"))
            {
                bool match = System.Text.RegularExpressions.Regex.IsMatch(format.ListString, "^\\w+\\.\\w+(?!\\.)$");
                if (paraStyle.ParagraphFormat.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevel2 || !match)
                {
                    multiError = true;
                }
            }
            else if (styleName.Equals("Heading 3"))
            {
                bool match = System.Text.RegularExpressions.Regex.IsMatch(format.ListString, "^\\w+\\.\\w+\\.\\w+(?!\\.)$");
                if (paraStyle.ParagraphFormat.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevel3 || !match)
                {
                    multiError = true;
                }
            }
            else if (styleName.Equals("Heading 4"))
            {
                bool match = System.Text.RegularExpressions.Regex.IsMatch(format.ListString, "^\\w+\\.\\w+\\.\\w+\\.\\w+(?!\\.)$");
                if (paraStyle.ParagraphFormat.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevel4 || !match)
                {
                    multiError = true;
                }
            }
            else if (styleName.Equals("Heading 5"))
            {
                bool match = System.Text.RegularExpressions.Regex.IsMatch(format.ListString, "^\\w+\\.\\w+\\.\\w+\\.\\w+\\.\\w+(?!\\.)$");
                if (paraStyle.ParagraphFormat.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevel5 || !match)
                {
                    multiError = true;
                }
            }
            else if (styleName.Equals("Heading 6"))
            {
                bool match = System.Text.RegularExpressions.Regex.IsMatch(format.ListString, "^\\w+\\.\\w+\\.\\w+\\.\\w+\\.\\w+\\.\\w+(?!\\.)$");
                if (paraStyle.ParagraphFormat.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevel6 || !match)
                {
                    multiError = true;
                }
            }

            if (styleName.Equals("Heading 7") || styleName.Equals("Heading 8") || styleName.Equals("Heading 9"))
            {
                if (!format.ListString.Equals(""))
                {
                    multiError = true;
                }
                String descrip = paraStyle.Description;
                String[] splitter = new String[1];
                splitter[0] = "Indent at:  0 cm";
                splitter = descrip.Split(splitter, StringSplitOptions.None);
                if (!(descrip.Length > 0))
                {
                    multiError = true;
                }
            }
        }

        /*
        *Retrieves the style object based on a number of strings in a string array
        */
        private Style[] getStyle(String[] names)
        {
            Style[] styles = new Style[names.Length];
            int stylesCounter = 0;
            foreach (Style s in doc.Styles)
            {

                for (int i = 0; i < names.Length; i++)
                {

                    if (s.NameLocal.Equals(names[i]))
                    {
                        styles[stylesCounter] = s;
                        stylesCounter++;
                        break;
                    }

                }
            }

            return styles;
        }

        /*
        * A method that will check a given paragraph matches certain attributes associated with the style. 
        */
        private bool checkParagraphStyle(Word.Paragraph p, Style actualStyle)
        {
            Style paraStyle = p.get_Style();
            String descrip = actualStyle.NameLocal;
            String[] splitter = new String[1];
            splitter[0] = "TOC";
            splitter = descrip.Split(splitter, StringSplitOptions.None);

            Word.Range theRange = p.Range;
            int numberOfTables = p.Range.Tables.Count;

            bool isEndnote = inEndnotes(p);//is this paragraph part of the endnotes

            Word.Find finder = p.Range.Find;
            finder.set_Style(hyperStyle);
            bool foundHyper = finder.Execute();
            var hyperParent = finder.Parent;
            String hyperlinkText = hyperParent.Text;

            bool foundQuoteChar = false;

            if (quotechar.NameLocal.Equals("Quote Char"))
            {
                theRange.Find.ClearFormatting();
                theRange.Find.set_Style(quotechar);
                foundQuoteChar = theRange.Find.Execute();
            }


            bool hasStrong = false;
            theRange.Find.ClearFormatting();
            theRange.Find.set_Style(strong);//changed to strong style of this class
            hasStrong = theRange.Find.Execute();

            bool hasEmphasis = false;
            theRange.Find.ClearFormatting();
            theRange.Find.set_Style(emphasis);//changed to emphasis of this class
            hasEmphasis = theRange.Find.Execute();

            if (!(splitter.Length > 1) && numberOfTables < 1 && !isEndnote && !paraStyle.NameLocal.Equals("Table of Figures") && !foundQuoteChar && !hasStrong && !hasEmphasis && !p.Range.Text.Equals("\r"))
            {

                if (foundHyper)//found a hyperlink
                {
                    finder = p.Range.Find;
                    finder.ClearFormatting();
                    finder.Font.Underline = actualStyle.Font.Underline;
                    bool gotUnderline = finder.Execute();
                    String underLinetext = finder.Parent.Text;//underline

                    finder = p.Range.Find;
                    finder.ClearFormatting();
                    finder.Font.Size = actualStyle.Font.Size;
                    bool gotFont = finder.Execute();
                    String sizeText = finder.Parent.Text;//fontsize

                    finder = p.Range.Find;
                    finder.ClearFormatting();
                    finder.Font.ColorIndex = actualStyle.Font.ColorIndex;
                    bool gotColor = finder.Execute();
                    String textColor = finder.Parent.Text;//font color

                    bool textOK = false;
                    String[] splitFromHyper = new String[1];
                    splitFromHyper[0] = hyperlinkText;
                    splitFromHyper = p.Range.Text.Split(splitFromHyper, StringSplitOptions.None);

                    foreach (String s in splitFromHyper)
                    {
                        if (s.Equals(underLinetext) && s.Equals(sizeText) && s.Equals(textColor))
                        {
                            textOK = true;
                        }
                    }

                    if (!textOK)
                    {
                        paragraphStyles.Add(p.Range.Text);
                        return false;
                    }
                }
                else
                {
                    // underline check
                    String lineAsStirng = p.Range.Font.Underline.ToString();
                    bool undefinedCheck = lineAsStirng.Equals("9999999");
                    if (p.Range.Font.Underline != actualStyle.Font.Underline || undefinedCheck)
                    {
                        if (!p.Range.Text.Equals("\r\a") && !p.Range.Text.Equals("\f") && !p.Range.Text.Equals("\r"))
                        {
                            paragraphStyles.Add(p.Range.Text);
                            return false;
                        }
                    }

                    //font size check
                    if (p.Range.Font.Size != actualStyle.Font.Size || p.Range.Font.Size == 9999999)
                    {
                        if (!p.Range.Text.Equals("\r\a") && !p.Range.Text.Equals("\f") && !p.Range.Text.Equals("\r"))
                        {

                            paragraphStyles.Add(p.Range.Text);
                            return false;
                        }
                    }

                    if (p.Range.Font.Color != actualStyle.Font.Color) //&& p.Range.Font.ColorIndex != -16777216)
                    {
                        if (!p.Range.Text.Equals("\r\a") && !p.Range.Text.Equals("\f") && !p.Range.Text.Equals("\r"))
                        {

                            paragraphStyles.Add(p.Range.Text);
                            return false;
                        }
                    }
                }


                //font face check
                if (!p.Range.Font.Name.Equals(actualStyle.Font.Name))
                {
                    if (!p.Range.Text.Equals("\r\a") && !p.Range.Text.Equals("\f") && !p.Range.Text.Equals("\r"))
                    {
                        paragraphStyles.Add(p.Range.Text);
                        return false;
                    }
                }


                //ensures alignment are the same
                if (p.Format.Alignment != actualStyle.ParagraphFormat.Alignment)
                {
                    if (!p.Range.Text.Equals("\r\a") && !p.Range.Text.Equals("\f") && !p.Range.Text.Equals("\r"))
                    {
                        paragraphStyles.Add(p.Range.Text);
                        return false;
                    }
                }

                if (p.Format.KeepWithNext != actualStyle.ParagraphFormat.KeepWithNext)
                {
                    paragraphStyles.Add(p.Range.Text);
                    return false;
                }

                //Uncomment this if you ever want to check for bold and italic has been applied to paragraphs
                //bold check
                //if (p.Range.Font.Bold != actualStyle.Font.Bold || p.Range.Font.Bold == 9999999)
                //{
                //    if(!p.Range.Text.Equals("\r\a"))
                //    {
                //    paragraphStyles.Add(p.Range.Text);
                //    return false;
                //    }
                //}

                //italic check
                //if (p.Range.Font.Italic != actualStyle.Font.Italic || p.Range.Font.Italic == 9999999)
                //{
                //    if (!p.Range.Text.Equals("\r\a"))
                //    {
                //        paragraphStyles.Add(p.Range.Text);
                //        return false;
                //    }
                //}

            }

            return true;
        }

        /*
         * A method that will check the carriage returns used in spacing method
         * */
        private bool checkCarriages(Word.Paragraph para)
        {
            bool carriageOk = true;
            if (para.Range.Text.Equals("\r"))
            {
                carriageCounter++;
                //if (carriageCounter == 2)
                //{

                //}
            }
            else
            {
                carriageCounter = 0;
            }

            if (carriageCounter == 2)
            {
                carriageOk = false;
                //carriageCounter = 0;
                //carriageOccasions = 0;
                //carriageCounter = 0;
            }

            return carriageOk;

        }

        /*
         * Checks that each heading style has been numbered. 
         */
        public bool headingsNumberedStyleTest(bool runOn, bool runError)
        {

            if (noMutli)
            {
                if (runOn)
                {
                    
                }
                return false;
            }
            else
            {
                if (multiError)
                {
                    if (runError)
                    {
                       
                    }
                    return false;
                }
            }
            return true;
        }

        public List<String> getSubtitleQuotes()
        {
            return this.subtitleStyleQuotes;
        }

        public List<Word.Style> getStylesInDoc()
        {
            return this.stylesInDoc;
        }

        public int getTitleCount()
        {
            return this.titleCount;
        }

        public bool titleusedThree()
        {
            return this.titleUsedThree;
        }

        public List<String> getCharacterQuotes()
        {
            return this.characterStyleQuotes;
        }

        public List<String> getFootNotes()
        {
            return this.footNoteStyleQuotes;
        }

        public int getDoubleCarriage()
        {
            return this.countDoubleCarriage;
        }

        /*
         * A style to test if the reference style has been used at somepoint 
         */
        public bool referenceStyleTest()
        {
            if (!refStyleUsed && endNoteResult == null)
            {
                return false;
            }
            return true;
        }

        public bool columnsTest()
        {
            if (this.columnsError)
            {
                return false;
            }
            return true;
        }

        public double getTop()
        {
            return this.topMargin;
        }

        public double getLeft()
        {
            return this.leftMargin;
        }

        public double getRight()
        {
            return this.rightMargin;
        }

        public double getBottom()
        {
            return this.bottomMargin;
        }

        public List<string> getNoExplainCaption()
        {
            return this.noExplanCaptionText;
        }

        public List<string> getCaptionObjects()
        {
            return this.captionWithObjects;
        }

        public List<string> getAllCaption()
        {
            return this.allCaptionText;
        }

        public int getReflink()
        {
            return this.refLinkBroken;
        }

        public int getNumberOfFigureCaps(){
            return this.numOfFigureCaps;
        }

        public int getNumberOfTableCaps(){
            return this.numOfTableCaps;
        }

        public int getNumberOfImages(){
            return this.numberOfImages;
        }

        public int getNumberOfTables(){
            return this.numberOfTables;
        }

        public int getNumberOfVrossRef(){
            return this.numOfCrossRef;
        }


        public void runDependencies()
        {
            foreach (Word.Field f in doc.Fields)
            {
                if (f.Type == Word.WdFieldType.wdFieldSequence)
                {
                    Word.Range range = f.Code;
                    Style captionStyle = range.get_Style();
                    String capString = range.Text;

                    bool isMatchTable = System.Text.RegularExpressions.Regex.IsMatch(capString.ToLower(), "table");
                    bool isMatchFigure = System.Text.RegularExpressions.Regex.IsMatch(capString.ToLower(), "figure");
                    
                    if (isMatchTable)
                    {
                        numOfTableCaps++;
                    }
                    else if (isMatchFigure)
                    {
                        numOfFigureCaps++;
                    }
                }

                if (f.Type == Word.WdFieldType.wdFieldRef)
                {
                    numOfCrossRef++;
                }
            }

            foreach (Word.InlineShape shape in doc.InlineShapes)
            {
                if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture || shape.Type == Word.WdInlineShapeType.wdInlineShapeChart)
                {
                    numberOfImages++;
                }
            }

            foreach (Word.Shape shape in doc.Shapes)
            {
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || shape.Type == Microsoft.Office.Core.MsoShapeType.msoChart || shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                {
                    numberOfImages++;

                }
                else if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                {
                    Word.Range contrange = shape.TextFrame.ContainingRange;
                    foreach (Word.Field field in contrange.Fields)
                    {
                        if (field.Type == Word.WdFieldType.wdFieldSequence)
                        {
                            Word.Range range = field.Code;
                            Style captionStyle = range.get_Style();
                            String capString = range.Text;

                            bool isMatchTable = System.Text.RegularExpressions.Regex.IsMatch(capString.ToLower(), "table");
                            bool isMatchFigure = System.Text.RegularExpressions.Regex.IsMatch(capString.ToLower(), "figure");

                            if (isMatchTable)
                            {
                                numOfTableCaps++;
                            }
                            else if (isMatchFigure)
                            {
                                numOfFigureCaps++;
                            }

                            Word.Range prevRange = contrange.Previous();
                            bool hasShapes = contrange.InlineShapes.Count > 0;//checks if has an inline image 
                            bool tableCount = contrange.Tables.Count > 0;// check for tables
                            String[] paraMark = contrange.Text.Split('\r');
                            bool isParaMark = paraMark[0].Equals("");

                            if (hasShapes || tableCount || isParaMark)
                            {
                                captionWithObjects.Add(contrange.Text);
                            }

                            String[] removeCarriage = contrange.Text.Split('\r');
                            // bool isMatch = System.Text.RegularExpressions.Regex.IsMatch(removeCarriage[0], "\\w*\\s*\\d+.{1}\\w+");
                            bool isMatch = System.Text.RegularExpressions.Regex.IsMatch(removeCarriage[0], "\\d.{1}\\D+|\\d.\\d\\D.{1}\\w+");
                            // Debug.Write(removeCarriage[0]); 
                            if (!isMatch)
                            {
                                noExplanCaptionText.Add(contrange.Text);//seperate list for captions that have no explanation 
                            }

                            allCaptionText.Add(contrange.Text);//add all captions to a list

                        }
                    }
                }
            }

            numberOfTables = doc.Tables.Count;

           

            //get some specfic styles used later in application 
            hyperStyle = getStyle("Hyperlink");
            quoteStyle = getStyle("Quote");
            strongStyle = getStyle("Strong");
            emphStyle = getStyle("Emphasis");
            quotechar = quoteStyle.get_LinkStyle();

            iterateOverPara();//will begin iteration over all paragraphs of the document 

        }

        public bool footerMEDITtest()
        {

            foreach (Word.Section s in doc.Sections)
            {
                foreach (Word.HeaderFooter foot in s.Footers)
                {
                    String text = foot.Range.Text;
                    bool match = System.Text.RegularExpressions.Regex.IsMatch(text, "MedIT\\.Exercises@ncl.ac.uk");
                    if (match)
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        public bool getMatchTempti()
        {
            return this.matchTempTip;
        }

        public bool getPlagerismStatement()
        {
            return this.matchPlagState;
        }

        public int getPageBreaks()
        {
            return this.countPageBreaks;
        }

        public float getBPBAny()
        {
            return this.countNPBAny;
        }

        public bool headingOrderTest()
        {
            if (headingOrderError)
            {
                return false;
            }
            return true;
        }

        public bool paragraphStyleTest(int instances)
        {
            int intstancesMinusChar = instances - characterStyleQuotes.Count;
            if (intstancesMinusChar >= 0)
            {
                if (paragraphStyles.Count >= instances)
                {

                    return false;
                   
                }
            }
            return true;
        }

        public int getNoOfShiftEnters()
        {
            return this.shiftEnters.Count;
        }

        public int getSpaceMiddle()
        {
            return this.spaceMiddle.Count;
        }
        public int getSpaceStart()
        {
            return this.spaceStart.Count;
        }
        public int getTabConsec()
        {
            return this.tabsConsec.Count;
        }
        public int getTaStart()
        {
            return this.tabsStart.Count;
        }
        
        public int getSingleCarriage()
        {
            return this.countSingleCarriage;
        }
        public List<string> getNormalwebQuotes()
        {
            return this.normalWebStyleQuotes;
        }

        public bool runTwoNext()
        {
            if (this.continousTwice)
            {
                return false;
            }
            return true;
        }
        public bool runInMiddle()
        {
            if (this.continiousInMiddle)
            {
                return false;
            }
            return true;
        }


    }


   }


