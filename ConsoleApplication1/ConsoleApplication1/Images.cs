using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
namespace ConsoleApplication1
{
    class Images:Styles
    {
        private int numberOfImages;
        private int numOfFigureCaps;


        public Images(Word.Document doc, Word.Application app,int number_of_images,int number_of_figures):base(doc,app)
        {
            this.numberOfImages = number_of_images;
            this.numOfFigureCaps = number_of_figures;
        }

        /*
         * Internal method used with imagesTest to check each image for certain attriburtes
         */
        private void shapeTests( Word.InlineShape shape, ref bool runTextWrap, ref bool runScale, ref bool runWidthHeight,
            ref bool runCrop, ref bool runImagePos,
            ref bool runImageCount, ref bool runImagesStyle, ref bool runPicturesStyle, int numberOfImages)
        {
            float height = shape.ScaleHeight;
            float width = shape.ScaleWidth;
            Word.Style theStyle = shape.Range.get_Style();

            if (!theStyle.NameLocal.Equals("Images"))
            {
                if (runImagesStyle)
                {
                    runImagesStyle = false;
                }
            }

            if (!theStyle.NameLocal.Equals("Picture"))
            {
                if (runPicturesStyle)
                {
                    runPicturesStyle = false;
                }
            }

            if ((height > 130 || height < 70 || width < 70 || width > 130) && height != 0 && width != 0)
            {
                if (runScale)
                {
                    runScale = false;
                }
            }

            if (height != width)
            {
                if (runWidthHeight)
                {
                    runWidthHeight = false;
                }

            }

            if (shape.PictureFormat.CropBottom > 0 || shape.PictureFormat.CropLeft > 0 || shape.PictureFormat.CropRight > 0 || shape.PictureFormat.CropTop > 0)
            {
                if (runCrop)
                {
                    runCrop = false;
                }
            }

            if (!imageHasCaption(shape))
            {
                if (runImagePos)
                {
                    runImagePos = false;
                }
            }

        }



        //Test if an image has a caption preceeding it 
        public bool imageHasCaption(Word.InlineShape shape)
        {
            Word.Paragraph p = shape.Range.Paragraphs.First;
            Word.Paragraph paraBefore = p.Previous();
            if (paraBefore != null)
            {
                foreach (Word.Field f in paraBefore.Range.Fields)
                {
                    if (f.Type == Word.WdFieldType.wdFieldSequence)
                    {
                        return true;
                    }
                }
            }

            if (paraBefore != null)
            {
                Word.Paragraph paraBefore2 = paraBefore.Previous();
                if (paraBefore2 != null)
                {
                    foreach (Word.Field f in paraBefore2.Range.Fields)
                    {
                        if (f.Type == Word.WdFieldType.wdFieldSequence)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        /*
         * Checks the images in a document
         */
        public bool imagesTest(bool runTextWrap, bool runScale, bool runWidthHeight, bool runCrop, bool runImagePos, bool runImageCount, bool runImageStyle, bool runPicturesStyle)
        {
            bool imageOK = true;
            bool groupHasCap = false;
            //FeedbackInfo feed = null;
            //float markTemp = 4 * numberOfImages;

            foreach (Word.Shape shape in app.ActiveDocument.Shapes)
            {
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                {
                    Word.GroupShapes groupShapes = shape.GroupItems;
                    for (int i = 1; i <= groupShapes.Count; i++)
                    {
                        if (groupShapes[i].Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                        {
                            Word.Paragraph paraBefore = shape.Anchor.Paragraphs.First.Previous();
                            if (paraBefore != null)
                            {
                                foreach (Word.Field f in paraBefore.Range.Fields)
                                {
                                    if (f.Type == Word.WdFieldType.wdFieldSequence)
                                    {
                                        groupHasCap = true;
                                    }
                                }
                            }

                            if (!groupHasCap)
                            {
                                if (paraBefore != null)
                                {
                                    Word.Paragraph paraBefore2 = paraBefore.Previous();
                                    if (paraBefore2 != null)
                                    {
                                        foreach (Word.Field f in paraBefore2.Range.Fields)
                                        {
                                            if (f.Type == Word.WdFieldType.wdFieldSequence)
                                            {
                                                groupHasCap = true;
                                            }
                                        }
                                    }
                                }
                            }

                            if (!groupHasCap)
                            {
                                if (runImagePos)
                                {
                                    runImagePos = false;
                                }

                            }

                            if (groupShapes[i].PictureFormat.CropBottom > 0 || groupShapes[i].PictureFormat.CropLeft > 0 || groupShapes[i].PictureFormat.CropRight > 0 || groupShapes[i].PictureFormat.CropTop > 0)
                            {
                                if (runCrop)
                                {
                                    runCrop = false;
                                }
                            }
                        }
                    }
                }
            }


            //Run shape test on all in line shapes
            foreach (Word.InlineShape shape in doc.InlineShapes)
            {
                if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture)
                {
                    shapeTests(shape, ref runTextWrap, ref runScale, ref runWidthHeight, ref runCrop, ref runImagePos,
                        ref runImageCount, ref runImageStyle, ref runPicturesStyle, numberOfImages);
                }
                else if (shape.Type == Word.WdInlineShapeType.wdInlineShapeChart)
                {
                    if (!imageHasCaption(shape))
                    {
                        
                        if (runImagePos)
                        {
                     
                            runImagePos = false;
                        }
                     
                    }

                }
            }

            if (runImageCount)
            {
                if (numberOfImages > numOfFigureCaps)
                {
                    imageOK = false;

                    runImageCount = false;
                }
            }

            if (runTextWrap)
            {
                //Run shape tests on images where text wrap is not in line
                foreach (Word.Shape s in doc.Shapes)
                {
                    Word.Shape s3 = s;
                    if (s.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                    {
                        Word.InlineShape s2 = s3.ConvertToInlineShape();//must convert to in line first
                        Word.WdInlineShapeType type = s2.Type;
                        if (type == Word.WdInlineShapeType.wdInlineShapePicture)
                        {
                            imageOK = false;
                            runTextWrap = false;

                            shapeTests(s2, ref runTextWrap, ref runScale, ref runWidthHeight, ref runCrop, ref runImagePos,
                        ref runImageCount, ref runImageStyle, ref runPicturesStyle, numberOfImages);
                        }
                    }
                    else if (s.Type == Microsoft.Office.Core.MsoShapeType.msoChart)
                    {
                        Word.InlineShape s2 = s3.ConvertToInlineShape();//must convert to in line first
                        runTextWrap = false;
                        if (!imageHasCaption(s2))
                        {
                            if (runImagePos)
                            {
                                runImagePos = false;
                            }
                        }
                    }
                }
            }

            return imageOK;
        }
    }
}
