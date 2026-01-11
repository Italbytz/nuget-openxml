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

            // Extract text from slide
            var slideText = ExtractTextFromSlide(slidePart);
            sb.AppendLine(slideText);
            sb.AppendLine();

            // If slide contains images, add them in a separate column
            if (!slidePart.ImageParts.Any()) continue;
            sb.AppendLine(":::");
            sb.AppendLine("::: {.column width=\"50%\"}");
            foreach (var imagePart in slidePart.ImageParts)
            {
                var imageDirectory = Path.Combine(outputDirectory, "images");
                if (!Directory.Exists(imageDirectory))
                    Directory.CreateDirectory(imageDirectory);
                var imageFileName = Path.GetFileName(imagePart.Uri.ToString());
                // ToDo: Handle duplicate image names
                /*var imageFileName =
                    $"{Guid.NewGuid()}{Path.GetExtension(imagePart.Uri.ToString())}";*/
                var imagePath =
                    Path.Combine(imageDirectory, imageFileName);

                // Save image to output directory
                using (var imageStream = imagePart.GetStream())
                using (var fileStream = File.Create(imagePath))
                {
                    imageStream.CopyTo(fileStream);
                }

                // Add image markdown
                sb.AppendLine($"![](images/{imageFileName})");
                sb.AppendLine();
            }

            sb.AppendLine(":::");
            sb.AppendLine("::::");
            sb.AppendLine();
        }

        var outputPath = Path.Combine(outputDirectory, fileName);

        File.WriteAllText(outputPath, sb.ToString());
    }


    private static string ExtractTextFromSlide(SlidePart slidePart)
    {
        var sb = new StringBuilder();

        var shapes = slidePart.Slide.Descendants<Shape>();
        foreach (var shape in shapes)
        {
            if (shape.IsFooter() || shape.IsHeader())
                continue;

            if (shape.IsTitle())
            {
                sb.AppendLine("## " + shape.TextBody?.InnerText.Trim());
                sb.AppendLine();

                // Check if slide contains images
                if (slidePart.ImageParts.Any())
                {
                    sb.AppendLine(":::: {.columns}");
                    sb.AppendLine();
                    sb.AppendLine("::: {.column width=\"50%\"}");
                }

                continue;
            }

            var textBody = shape.TextBody;
            if (textBody != null)
                foreach (var paragraph in textBody.Elements<A.Paragraph>())
                {
                    var paragraphProperties = paragraph.ParagraphProperties;
                    var level = paragraphProperties?.Level?.Value ?? 0;
                    var isBulleted =
                        paragraphProperties?.GetFirstChild<A.BulletFont>() !=
                        null ||
                        paragraphProperties
                            ?.GetFirstChild<A.AutoNumberedBullet>() == null;
                    var isNumbered =
                        paragraphProperties
                            ?.GetFirstChild<A.AutoNumberedBullet>() != null;
                    var indent = new string(' ', level * 2);
                    foreach (var run in paragraph.Elements<A.Run>())
                    {
                        
                        var text = run.Text?.Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            if (isBulleted)
                                sb.Append(
                                    $"{indent}- {text}");
                            else if (isNumbered)
                                sb.Append(
                                    $"{indent}1. {text}");
                            else
                                sb.Append(text);
                        }
                    }

                    sb.AppendLine();
                }
        }

        return sb.ToString().Trim();
    }
}