using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OXML.Library
{
    public class Helper_PPT
    {

        public int PPTGetSlideCount(string fileName)
        {
            return PPTGetSlideCount(fileName, true);
        }
        public int PPTGetSlideCount(string fileName, bool includeHidden)
        {
            int slidesCount = 0;

            using (PresentationDocument presentationDocument = PresentationDocument.Open(fileName, false))
            {
                //Get the presentation part of the document
                PresentationPart presentationPart = presentationDocument.PresentationPart;
                if (presentationPart != null)
                {
                    if (includeHidden)
                    {
                        slidesCount = presentationPart.SlideParts.Count();
                        //slidesCount = presentationPart.GetPartsOfType<SlidePart>().Count();
                    }
                    else
                    {
                        var slides = presentationPart.SlideParts.
                            Where((s) => (s.Slide != null) && ((s.Slide.Show == null) || (s.Slide.Show.HasValue && s.Slide.Show.Value)));
                        slidesCount = slides.Count();
                    }
                }
            }
            return slidesCount;
        }
    }
}
