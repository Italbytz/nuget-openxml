using DocumentFormat.OpenXml.Packaging;

namespace Italbytz.OpenXml;

public static class PresentationConverter
{
    public static void ConvertToQuartoMarkdown(string presentationFile,
        string destinationFile, string imageDirectory,
        string title = "Presentation",
        string? author = null)
    {
        using var presentationDocument =
            PresentationDocument.Open(presentationFile, false);
        presentationDocument.ExportToQuartoMarkdown(
            destinationFile, imageDirectory, title, author);
    }
}