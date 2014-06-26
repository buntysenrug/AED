using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word=Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class ShapeTest:Styles
    {
        private Word.InlineShape shape;
        private float width;
        private float height;
        public ShapeTest(Word.Document doc, Word.Application app, Word.InlineShape shape):base(doc,app)
        {
            this.shape = shape;
            this.width = shape.Width;
            this.height = shape.Height;
        }

        public bool runImagesStyle()
        {
            Word.Style style = shape.Range.get_Style();
            if (!style.NameLocal.Equals("Images"))
            {
                return false;
            }
            return true;
        }

        public bool runPicturesStyle()
        {

            Word.Style style = shape.Range.get_Style();
            if (!style.NameLocal.Equals("Picture"))
            {
                return false;
            }
            return true;
        }

        public bool runScale()
        {
            if ((height > 130 || height < 70 || width < 70 || width > 130) && height != 0 && width != 0)
            {
                return false;
            }
            return true;
        }

        public bool runWidthHeight()
        {
            if (height != width)
            {
                return false;
            }
            return true;
        }

        public bool runCrop()
        {
            if (shape.PictureFormat.CropBottom > 0 || shape.PictureFormat.CropLeft > 0 || shape.PictureFormat.CropRight > 0 ||
                shape.PictureFormat.CropTop > 0)
            {
                return false;
            }
            return true;
        }

        //Test if an image has a caption preceeding it 
        public bool imageHasCaption()
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

        public bool runImagePos()
        {
            if (!imageHasCaption())
            {
                return false;
            }
            return true;
        }



    }
}
