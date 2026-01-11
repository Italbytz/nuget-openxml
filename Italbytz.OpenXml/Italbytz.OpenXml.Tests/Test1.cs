using System.Text;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;
using Text = DocumentFormat.OpenXml.Drawing.Text;

namespace Italbytz.OpenXml.Tests;

[TestClass]
public sealed class Test1
{
    [TestMethod]
    public void TestMethod4()
    {
        var path = "/Users/nunkesser/repos/work/ppt/PM-Folien/LE01.pptx";
        using var presentationDocument =
            PresentationDocument.Open(path, false);
        presentationDocument.ExportToQuartoMarkdown(
            "/Users/nunkesser/repos/work/ppt/PM-Folien",
            "LE01.qmd");
    }

    [TestMethod]
    public void TestMethod3()
    {
        var path = "/Users/nunkesser/repos/work/ppt/PM-Folien/LE01.pptx";
        var titles = GetSlideTitles(path);
        for (var i = 0; i < titles.Count; i++)
            Console.WriteLine($"Title of slide #{i + 1} is: {titles[i]}");
    }

    [TestMethod]
    public void TestMethod2()
    {
        var path = "/Users/nunkesser/repos/work/ppt/PM-Folien/LE01.pptx";
        var numberOfSlides = CountSlides(path);
        Console.WriteLine($"Number of slides = {numberOfSlides}");

        for (var i = 0; i < numberOfSlides; i++)
        {
            GetSlideIdAndText(out var text, path, i);
            Console.WriteLine($"Side #{i + 1} contains: {text}");
        }
    }

    [TestMethod]
    public void TestMethod1()
    {
        using (var presentationDocument =
               PresentationDocument.Open(
                   "/Users/nunkesser/repos/work/ppt/PM-Folien/LE01.pptx",
                   false))
        {
            // Get the presentation part from the presentation document.
            var presentationPart = presentationDocument.PresentationPart;
            var presentation = presentationPart.Presentation;

            // Create a StringBuilder for the markdown output
            var markdown = new StringBuilder();
            markdown.AppendLine("---");

            // Iterate through each slide
            foreach (var slideId in presentation.SlideIdList.ChildElements
                         .OfType<SlideId>())
            {
                var slidePart =
                    (SlidePart)presentationPart.GetPartById(
                        slideId.RelationshipId);

                // Get slide title (usually first shape with text)
                var titleShape = slidePart.Slide.Descendants<Shape>()
                    .FirstOrDefault(sp =>
                        sp.ShapeProperties?.Transform2D != null);

                if (titleShape != null)
                {
                    var titleText = titleShape.Descendants<Text>()
                        .Select(t => t.Text)
                        .Aggregate("", (current, text) => current + text);

                    markdown.AppendLine($"## {titleText}");
                    markdown.AppendLine();
                }

                // Get all text content from other shapes
                foreach (var shape in slidePart.Slide.Descendants<Shape>()
                             .Skip(1))
                {
                    var shapeText = shape.Descendants<Text>()
                        .Select(t => t.Text)
                        .Aggregate("", (current, text) => current + text);

                    if (!string.IsNullOrWhiteSpace(shapeText))
                        markdown.AppendLine($"- {shapeText}");
                }

                markdown.AppendLine();
                markdown.AppendLine("---");
                markdown.AppendLine();
            }

            // Save to file
            File.WriteAllText(
                "/Users/nunkesser/repos/work/ppt/PM-Folien/LE01.qmd",
                markdown.ToString());
            Console.WriteLine("Markdown file created successfully!");
        }
    }

    private static int CountSlides(string presentationFile)
    {
        // <Snippet1>
        // Open the presentation as read-only.
        using (var presentationDocument =
               PresentationDocument.Open(presentationFile, false))
            // </Snippet1>
        {
            // Pass the presentation to the next CountSlides method
            // and return the slide count.
            return CountSlidesFromPresentation(presentationDocument);
        }
    }

// Count the slides in the presentation.
    private static int CountSlidesFromPresentation(
        PresentationDocument presentationDocument)
    {
        // Check for a null document object.
        if (presentationDocument is null)
            throw new ArgumentNullException("presentationDocument");

        var slidesCount = 0;

        // Get the presentation part of document.
        var presentationPart = presentationDocument.PresentationPart;
        // Get the slide count from the SlideParts.
        if (presentationPart is not null)
            slidesCount = presentationPart.SlideParts.Count();

        // Return the slide count to the previous method.
        return slidesCount;
    }

    private static void GetSlideIdAndText(out string sldText, string docName,
        int index)
    {
        using (var ppt = PresentationDocument.Open(docName, false))
        {
            // Get the relationship ID of the first slide.
            var part = ppt.PresentationPart;
            var slideIds = part?.Presentation?.SlideIdList?.ChildElements ??
                           default;

            if (part is null || slideIds.Count == 0)
            {
                sldText = "";
                return;
            }

            string? relId = ((SlideId)slideIds[index]).RelationshipId;

            if (relId is null)
            {
                sldText = "";
                return;
            }

            // Get the slide part from the relationship ID.
            var slide = (SlidePart)part.GetPartById(relId);

            // Build a StringBuilder object.
            var paragraphText = new StringBuilder();

            // Get the inner text of the slide:
            var texts = slide.Slide.Descendants<Text>();
            foreach (var text in texts) paragraphText.Append(text.Text);
            sldText = paragraphText.ToString();
        }
    }

    // Get a list of the titles of all the slides in the presentation.
    private static IList<string> GetSlideTitles(string presentationFile)
    {
        // <Snippet1>
        // Open the presentation as read-only.
        using (var presentationDocument =
               PresentationDocument.Open(presentationFile, false))
            // </Snippet1>
        {
            var titles = GetSlideTitlesFromPresentation(presentationDocument);

            return (IList<string>)(titles ?? Enumerable.Empty<string>());
        }
    }

// Get a list of the titles of all the slides in the presentation.
    private static IList<string>? GetSlideTitlesFromPresentation(
        PresentationDocument presentationDocument)
    {
        // Get a PresentationPart object from the PresentationDocument object.
        var presentationPart = presentationDocument.PresentationPart;

        if (presentationPart is not null &&
            presentationPart.Presentation is not null)
        {
            // Get a Presentation object from the PresentationPart object.
            var presentation = presentationPart.Presentation;

            if (presentation.SlideIdList is not null)
            {
                var titlesList = new List<string>();

                // Get the title of each slide in the slide order.
                foreach (var slideId in presentation.SlideIdList
                             .Elements<SlideId>())
                {
                    if (slideId.RelationshipId is null) continue;

                    var slidePart =
                        (SlidePart)presentationPart.GetPartById(
                            slideId.RelationshipId!);

                    // Get the slide title.
                    var title = GetSlideTitle(slidePart);

                    // An empty title can also be added.
                    titlesList.Add(title);
                }

                return titlesList;
            }
        }

        return null;
    }

// Get the title string of the slide.
    private static string GetSlideTitle(SlidePart slidePart)
    {
        if (slidePart is null)
            throw new ArgumentNullException("presentationDocument");

        // Declare a paragraph separator.
        string? paragraphSeparator = null;

        if (slidePart.Slide is not null)
        {
            // Find all the title shapes.
            var shapes = from shape in slidePart.Slide.Descendants<Shape>()
                where IsTitleShape(shape)
                select shape;

            var paragraphText = new StringBuilder();

            foreach (var shape in shapes)
            {
                var paragraphs = shape.TextBody?.Descendants<Paragraph>();
                if (paragraphs is null) continue;

                // Get the text in each paragraph in this shape.
                foreach (var paragraph in paragraphs)
                {
                    // Add a line break.
                    paragraphText.Append(paragraphSeparator);

                    foreach (var text in paragraph.Descendants<Text>())
                        paragraphText.Append(text.Text);

                    paragraphSeparator = "\n";
                }
            }

            return paragraphText.ToString();
        }

        return string.Empty;
    }

// Determines whether the shape is a title shape.
    private static bool IsTitleShape(Shape shape)
    {
        var placeholderShape = shape.NonVisualShapeProperties
            ?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();

        if (placeholderShape is not null && placeholderShape.Type is not null &&
            placeholderShape.Type.HasValue)
            return placeholderShape.Type == PlaceholderValues.Title ||
                   placeholderShape.Type == PlaceholderValues.CenteredTitle;

        return false;
    }
}