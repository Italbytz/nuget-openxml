using DocumentFormat.OpenXml.Packaging;

namespace Italbytz.OpenXml.Tests;

[TestClass]
public sealed class Test1
{
    [TestMethod]
    public void TestMethod1()
    {
        using (PresentationDocument presentationDocument =
       PresentationDocument.Open("/Users/nunkesser/repos/work/ppt/PM-Folien/LE01.pptx", false))
{
    // Get the presentation part from the presentation document.
    var presentationPart = presentationDocument.PresentationPart;
    var presentation = presentationPart.Presentation;
    
    // Create a StringBuilder for the markdown output
    var markdown = new System.Text.StringBuilder();
    markdown.AppendLine("---");
    
    // Iterate through each slide
    foreach (var slideId in presentation.SlideIdList.ChildElements.OfType<DocumentFormat.OpenXml.Presentation.SlideId>())
    {
        var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
        
        // Get slide title (usually first shape with text)
        var titleShape = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>()
            .FirstOrDefault(sp => sp.ShapeProperties?.Transform2D != null);
        
        if (titleShape != null)
        {
            var titleText = titleShape.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                .Select(t => t.Text)
                .Aggregate("", (current, text) => current + text);
            
            markdown.AppendLine($"## {titleText}");
            markdown.AppendLine();
        }
        
        // Get all text content from other shapes
        foreach (var shape in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().Skip(1))
        {
            var shapeText = shape.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                .Select(t => t.Text)
                .Aggregate("", (current, text) => current + text);
            
            if (!string.IsNullOrWhiteSpace(shapeText))
            {
                markdown.AppendLine($"- {shapeText}");
            }
        }
        
        markdown.AppendLine();
        markdown.AppendLine("---");
        markdown.AppendLine();
    }
    
    // Save to file
    System.IO.File.WriteAllText("/Users/nunkesser/repos/work/ppt/PM-Folien/LE01.qmd", markdown.ToString());
    Console.WriteLine("Markdown file created successfully!");
}
    }
}