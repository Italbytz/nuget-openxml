using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace Italbytz.OpenXml;

using A = DocumentFormat.OpenXml.Drawing;

public static class PresentationDocumentExtensions
{
    public static void ExportToQuartoMarkdown(
        this PresentationDocument presentationDocument, string outputDirectory,
        string fileName = "presentation.qmd", string title = "Presentation",
        string? author = null)
    {
        ArgumentNullException.ThrowIfNull(presentationDocument);

        if (string.IsNullOrWhiteSpace(outputDirectory))
            throw new ArgumentException("Output path cannot be null or empty.",
                nameof(outputDirectory));

        var directory = Path.GetDirectoryName(outputDirectory);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            Directory.CreateDirectory(directory);

        var sb = new StringBuilder();

        // Add Quarto/Reveal.js YAML header
        sb.AppendLine("---");
        sb.AppendLine($"title: \"{title}\"");
        if (!string.IsNullOrWhiteSpace(author))
            sb.AppendLine($"author: \"{author}\"");
        sb.AppendLine("format: revealjs");
        sb.AppendLine("---");
        sb.AppendLine();
        sb.AppendLine("# ");
        sb.AppendLine();

        var presentationPart = presentationDocument.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList == null)
            return;

        // Iterate through slides
        foreach (var slideId in presentationPart.Presentation.SlideIdList
                     .Elements<SlideId>())
        {
            var slidePart =
                (SlidePart)presentationPart.GetPartById(
                    slideId.RelationshipId!);

            // Extract images from slide
            foreach (var imagePart in slidePart.ImageParts)
            {
                var imageDirectory = Path.Combine(outputDirectory, "images");
                if (!Directory.Exists(imageDirectory))
                    Directory.CreateDirectory(imageDirectory);
                var imageFileName = Path.GetFileName(imagePart.Uri.ToString());
                var imagePath = Path.Combine(imageDirectory, imageFileName);
                using var imageStream = imagePart.GetStream();
                using var fileStream = File.Create(imagePath);
                imageStream.CopyTo(fileStream);
                sb.AppendLine($"![](images/{imageFileName})");
                sb.AppendLine();
            }

            // Extract text from slide
            var slideText = ExtractTextFromSlide(slidePart);
            sb.AppendLine(slideText);
            sb.AppendLine();
        }

        var outputPath = Path.Combine(outputDirectory, fileName);

        File.WriteAllText(outputPath, sb.ToString());
    }

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


    private static bool IsFooterShape(Shape shape)
    {
        var placeholderShape = shape.NonVisualShapeProperties
            ?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();

        if (placeholderShape is not null && placeholderShape.Type is not null &&
            placeholderShape.Type.HasValue)
            return placeholderShape.Type == PlaceholderValues.Footer ||
                   placeholderShape.Type == PlaceholderValues.SlideNumber ||
                   placeholderShape.Type == PlaceholderValues.DateAndTime;

        return false;
    }

    private static string ExtractTextFromSlide(SlidePart slidePart)
    {
        var sb = new StringBuilder();

        var shapes = slidePart.Slide.Descendants<Shape>();
        foreach (var shape in shapes)
        {
            if (IsFooterShape(shape))
                continue;

            if (IsTitleShape(shape))
            {
                sb.AppendLine("## " + shape.TextBody?.InnerText.Trim());
                sb.AppendLine();
                continue;
            }

            var textBody = shape.TextBody;
            if (textBody != null)
                foreach (var paragraph in textBody.Elements<A.Paragraph>())
                {
                    foreach (var run in paragraph.Elements<A.Run>())
                    {
                        var text = run.Text?.Text;
                        if (!string.IsNullOrWhiteSpace(text)) sb.Append(text);
                    }

                    sb.AppendLine();
                }
        }

        return sb.ToString().Trim();
    }
}