using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using D = DocumentFormat.OpenXml.Drawing;
using Presentationnote = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using F = DocumentFormat.OpenXml.Drawing;



class pptxDoco
{
    private PresentationDocument doco;
    private int nTotalSlides; // -1 can either indicate error or just not having run the method to calculate this
    private String szExceptionDetails; //For storing the results of actual exceptions
    private String szErrorDetails; //For storing the results of errors

    //Constructors
    public pptxDoco()
            {
                doco = null;
                nTotalSlides = -1;
                szExceptionDetails = null;
                szErrorDetails = null;
            }
    public pptxDoco(String szPathToDoc)
            {
                try
                {
                    doco = PresentationDocument.Open(szPathToDoc, true);
                    Console.WriteLine("{0} has been opened", szPathToDoc);
                    szExceptionDetails = null;
                    szErrorDetails = null;
                    nTotalSlides = -1;
                }
                catch (System.Exception e)
                {
                    //Setting the doco to null and sending the details of the exception to the string for this
                    doco = null;
                    nTotalSlides = -1;
                    szErrorDetails = null;
                    szExceptionDetails = e.ToString();
                }

            }

    //Helper methods
    private static string GetSlideTitle(SlidePart slidePart)
            {
                if (slidePart == null)
                {
                    throw new ArgumentNullException("presentationDocument");
                }

                // Declare a paragraph separator.
                string paragraphSeparator = null;

                if (slidePart.Slide != null)
                {
                    // Find all the title shapes.
                    var shapes = from shape in slidePart.Slide.Descendants<Shape>()
                                 where IsTitleShape(shape)
                                 select shape;

                    StringBuilder paragraphText = new StringBuilder();

                    foreach (var shape in shapes)
                    {
                        // Get the text in each paragraph in this shape.
                        foreach (var paragraph in shape.TextBody.Descendants<D.Paragraph>())
                        {
                            // Add a line break.
                            paragraphText.Append(paragraphSeparator);

                            foreach (var text in paragraph.Descendants<D.Text>())
                            {
                                paragraphText.Append(text.Text);
                            }

                            paragraphSeparator = "\n";
                        }
                    }

                    return paragraphText.ToString();
                }

                return string.Empty;
            }
    private static bool IsTitleShape(Shape shape)
            {
                var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
                if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
                {
                    switch ((PlaceholderValues)placeholderShape.Type)
                    {
                        // Any title shape.
                        case PlaceholderValues.Title:

                        // A centered title.
                        case PlaceholderValues.CenteredTitle:
                            return true;

                        default:
                            return false;
                    }
                }
                return false;
            }
    private static string GetSlideLayoutType(SlideLayoutPart slideLayoutPart)
            {
                CommonSlideData slideData = slideLayoutPart.SlideLayout.CommonSlideData;

                // Remarks: If this is used in production code, check for a null reference.

                return slideData.Name;
            }

    //Display Methods
    public int getCount()
            {
                if (nTotalSlides == -1 && doco != null) { nTotalSlides = CountSlides(); }
                return nTotalSlides;
            }
    public String getError()
            {
                return szErrorDetails;
            }
    public String getException()
            {
                return szExceptionDetails;
            }

    //Methods that actually do useful things
    public int CountSlides()
            {
                //This would indicate this method has been run so let's just cut to the chase here

                if (nTotalSlides != -1)
                {
                    return nTotalSlides;
                }

                //Protecting against people who will try and run this against an empty class
                if (doco == null)
                {
                    szErrorDetails = "No Document has been specified in the constructor";
                    return nTotalSlides;
                }
                PresentationPart part = doco.PresentationPart;
                if (part == null)
                {
                    szErrorDetails = "Presentation included has no presentation parts!";
                    return nTotalSlides;
                }

                nTotalSlides = part.SlideParts.Count();
                return nTotalSlides;
            }
    public void getSlideIDAndText(out String szSlideTxt, int index)
            {
                //Initial Error checking and stuff
                int error_check = CountSlides();
                if (error_check == -1)
                {
                    szErrorDetails = "GetSlideIDAndText: This Slide Count is returning -1.....stop messing with us";
                    szSlideTxt = "Panic Panic this is not right!!!!! Crashing and Burning!!!!!!!!";
                    szErrorDetails = "Something went wrong with the Count Slides method...could have been: No doc in constructor or no presentation parts in the doc";
                    return;
                }

                //Time to get the relationship ID of the first slide
                PresentationPart part = doco.PresentationPart;
                if (part == null)
                {
                    szSlideTxt = "Panic Panic this is not right!!!!! Crashing and Burning!!!!!!!!....this is number 2 btw";
                    szErrorDetails = "Basically on this one the Presentation stored in this object doesn't have parts......like a head....or legs....or a heart";
                    return;
                }
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[index] as SlideId).RelationshipId;

                //Getting the Slide part from the relationship ID
                SlidePart slide = (SlidePart)part.GetPartById(relId);

                //Building a stringbuilder object
                StringBuilder paragraphText = new StringBuilder();

                //Get the inner text of the slide
                IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
                foreach (A.Text text in texts)
                {
                    paragraphText.Append(text.Text);
                }
                szSlideTxt = paragraphText.ToString();
            }
    public IList<string> getSlideTitles()
            {
                string title = null;
                if (doco == null) { throw new ArgumentNullException("Doco hasn't been initialized and it needs to be for this operation"); }
                //Grabbing the presentation part
                PresentationPart part = doco.PresentationPart;
                //Checking to make sure the part and presentation object in the part are not null
                if (part == null || part.Presentation == null)
                { throw new ArgumentNullException("Presentation or Presentation part is null"); }
                //Getting the presentation object from the presentation part
                Presentation present = part.Presentation;
                if (present.SlideIdList == null) { throw new ArgumentNullException("Danger Will Robinson! Slide ID List in the presentation is null"); }
                List<string> titlesList = new List<string>();
                //Getting the title of each slide in the slide order
                foreach (var slideId in present.SlideIdList.Elements<SlideId>())
                {
                    SlidePart slidePart = part.GetPartById(slideId.RelationshipId) as SlidePart;

                    //Get the slide title
                    title = GetSlideTitle(slidePart);

                    //If the title is empty it's all good
                    titlesList.Add(title);
                }
                return titlesList;
            }
    public void deleteSlide(int index)
            {
                if (nTotalSlides == -1) { CountSlides(); }
                if (index < 0 || index > nTotalSlides) { throw new ArgumentOutOfRangeException("The slide index provided is either too low or too high."); }
                if (doco == null) { throw new NullReferenceException("The PresentationDocument is set to null, this is a bad bad thing"); }
                //Grabbing the Presentation Part from doco
                PresentationPart part = doco.PresentationPart;
                //Testing for null in either part or part.presentation
                if (part == null || part.Presentation == null) { throw new NullReferenceException("The PresentationPart or the Presentation is not valid"); }
                Presentation pres = part.Presentation;
                //Grabbing the list of slide IDs in the presentation
                SlideIdList slideIdList = pres.SlideIdList;
                //Getting the Slide ID of the specified slide (specified by index)
                SlideId slideId = slideIdList.ChildElements[index] as SlideId;
                //Getting the relationship ID of the slide
                string slideRelId = slideId.RelationshipId;
                //Remove the slide form the slide list
                slideIdList.RemoveChild(slideId);
                //Removing references to the slide from all custom shows
                if (pres.CustomShowList != null)
                {
                    //Going through the custom shows
                    foreach (var customShow in pres.CustomShowList.Elements<CustomShow>())
                    {
                        if (customShow.SlideList != null)
                        {
                            //Declaring a link list of slide list entries
                            LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                            foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                            {
                                //Find the slide reference to remove from the custom show
                                if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                                {
                                    slideListEntries.AddLast(slideListEntry);
                                }
                            }
                            //Removing all refrences to the slide from the custom show
                            foreach (SlideListEntry slideListEntry in slideListEntries)
                            {
                                customShow.SlideList.RemoveChild(slideListEntry);
                            }
                        }
                    }
                }
                //Saving the modified presentation
                pres.Save();
                //Getting the slide part for the specified slide
                SlidePart slidePart = part.GetPartById(slideRelId) as SlidePart;
                //Delete the slide part
                part.DeletePart(slidePart);
                this.CountSlides();
            }
    public void changeTemplate(string templateName)
    {
        PresentationDocument templateDoco = PresentationDocument.Open(templateName, false);
        if (templateDoco == null || doco == null)
        {
            throw new ArgumentException("Initial Document or Presentation document aren't initialized");
        }
        //Getting the presentation part of the initial document
        PresentationPart part = doco.PresentationPart;
        //Grabbing the slide master part
        SlideMasterPart slideMasterPart = part.SlideMasterParts.ElementAt(0);
        string relationshipId = part.GetIdOfPart(slideMasterPart);
        //Grabbing the new slide master part
        SlideMasterPart newSlideMasterPart = templateDoco.PresentationPart.SlideMasterParts.ElementAt(0);
        //Remove the existing theme part
        part.DeletePart(part.ThemePart);
        //Remove old slide master part
        part.DeletePart(slideMasterPart);
        //import the new slide paster part, and reuse old relationship ID
        newSlideMasterPart = part.AddPart(newSlideMasterPart, relationshipId);
        //Change to the new theme part
        part.AddPart(newSlideMasterPart.ThemePart);
        Dictionary<string, SlideLayoutPart> newSlideLayouts = new Dictionary<string, SlideLayoutPart>();
        foreach (var slideLayoutPart in newSlideMasterPart.SlideLayoutParts)
        {
            newSlideLayouts.Add(GetSlideLayoutType(slideLayoutPart), slideLayoutPart);
        }
        string layoutType = null;
        SlideLayoutPart newLayoutPart = null;
        //Insert the code for the layout for this example
        string defaultLayoutType = "Title and Content";
        //Remove the slide layout relationship on all slides
        foreach (var slidePart in part.SlideParts)
        {
            layoutType = null;
            if (slidePart.SlideLayoutPart != null)
            {
                //Determine the slide layout type for each slide
                layoutType = GetSlideLayoutType(slidePart.SlideLayoutPart);
                //Delete old layout part
                slidePart.DeletePart(slidePart.SlideLayoutPart);
            }
            if (layoutType != null && newSlideLayouts.TryGetValue(layoutType, out newLayoutPart))
            {
                //Apply the new layout part
                slidePart.AddPart(newLayoutPart);
            }
            else
            {
                newLayoutPart = newSlideLayouts[defaultLayoutType];
                //Apply the new default layout part
                slidePart.AddPart(newLayoutPart);
             }
        }
        templateDoco.Close();
    }
    public void deleteNotes()
    {
        //I ripped this from: http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2011/08/16/screen-cast-remove-speaker-notes-from-an-open-xml-presentation.aspx
        //Error checking for doco being null goes here

        //Isolating out the slide part
        PresentationPart part = doco.PresentationPart;
        //Error checking for part being null goes here

        //Checking through the presentation parts
        foreach (var slide in part.SlideParts)
        {
            NotesSlidePart notes = slide.NotesSlidePart;
            if (notes!=null)
            {
                var sp = notes
                         .NotesSlide
                         .CommonSlideData
                         .ShapeTree
                         .Elements<Presentationnote.Shape>()
                         .FirstOrDefault(s => {
                             var nvSpPr = s.NonVisualShapeProperties;
                             if (nvSpPr != null)
                             {
                                 var nvPr = nvSpPr.ApplicationNonVisualDrawingProperties;
                                 if (nvPr != null)
                                 {
                                     var ph = nvPr.PlaceholderShape;
                                     if (ph != null)
                                         return ph.Type == PlaceholderValues.Body;
                                 }
                             }
                             return false;
                         });
                if (sp != null)
                {
                    var textBody = sp.TextBody;
                    if (textBody != null)
                    {
                        var firstParagraph = textBody
                            .Elements<F.Paragraph>()
                            .FirstOrDefault();
                        if (firstParagraph != null)
                        {
                            List<F.Paragraph> subsequentParagraphs = textBody
                                .Elements<F.Paragraph>()
                                .Skip(1)
                                .ToList();
                            firstParagraph.RemoveAllChildren();
                            foreach (var item in subsequentParagraphs)
                                item.Remove();
                        }
                    }
                }
            }
        }
    }
    public void closefile()
    {
        doco.Close();
    }
}
