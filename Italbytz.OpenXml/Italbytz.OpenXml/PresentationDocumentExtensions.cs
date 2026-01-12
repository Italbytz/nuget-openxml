using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace Italbytz.OpenXml;

using A = DocumentFormat.OpenXml.Drawing;

public static class PresentationDocumentExtensions
{
    public static void ExportToQuartoMarkdown(
        this PresentationDocument presentationDocument, string destinationFile,
        string imageDirectory, string title = "Presentation",
        string? author = null)
    {
        ArgumentNullException.ThrowIfNull(presentationDocument);

        if (string.IsNullOrWhiteSpace(destinationFile))
            throw new ArgumentException("Output path cannot be null or empty.",
                nameof(destinationFile));

        var directory = Path.GetDirectoryName(destinationFile);
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
                if (!Directory.Exists(imageDirectory))
                    Directory.CreateDirectory(imageDirectory);
                //var imageFileName = Path.GetFileName(imagePart.Uri.ToString());

                var imageFileName =
                    $"{Guid.NewGuid()}{Path.GetExtension(imagePart.Uri.ToString())}";
                var imagePath =
                    Path.Combine(imageDirectory, imageFileName);

                // Save image to output directory
                using (var imageStream = imagePart.GetStream())
                using (var fileStream = File.Create(imagePath))
                {
                    imageStream.CopyTo(fileStream);
                }

                // Add image markdown
                sb.AppendLine($"![]({imagePath})");
                sb.AppendLine();
            }

            sb.AppendLine(":::");
            sb.AppendLine("::::");
            sb.AppendLine();
        }

        File.WriteAllText(destinationFile, sb.ToString());
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
            {
                var autoNumber = 0;
                foreach (var paragraph in textBody.Elements<A.Paragraph>())
                {
                    var paragraphProperties = paragraph.ParagraphProperties;
                    var level = paragraphProperties?.Level?.Value ?? 0;
                    var isBulleted =
                        paragraphProperties
                            ?.GetFirstChild<A.NoBullet>() == null;
                    var isNumbered =
                        paragraphProperties
                            ?.GetFirstChild<A.AutoNumberedBullet>() != null;
                    var indent = new string(' ', level * 2);
                    var hasText = paragraph.Elements<A.Run>().Any(r =>
                        !string.IsNullOrWhiteSpace(r.Text?.Text));
                    if (hasText)
                    {
                        if (isNumbered)
                            sb.Append(indent + $"{++autoNumber}. ");
                        else if (isBulleted)
                            sb.Append(indent + "- ");
                        else
                            sb.Append(indent);
                    }

                    foreach (var run in paragraph.Elements<A.Run>())
                    {
                        var text = run.Text?.Text;
                        if (!string.IsNullOrWhiteSpace(text))
                            sb.Append(text);
                        else
                            sb.Append(" ");
                    }

                    if (!hasText) continue;
                    sb.AppendLine();
                    if (!isBulleted && !isNumbered)
                        sb.AppendLine();
                }
            }
        }

        return sb.ToString().Trim();
    }
}