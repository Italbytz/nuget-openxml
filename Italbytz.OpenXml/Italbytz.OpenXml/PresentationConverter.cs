using DocumentFormat.OpenXml.Packaging;

namespace Italbytz.OpenXml;

public static class PresentationConverter
{
    public static void ConvertToQuartoMarkdown(string presentationFile,
        string outputDirectory)
    {
        var fileName = Path.GetFileNameWithoutExtension(presentationFile) +
                       ".qmd";
        using var presentationDocument =
            PresentationDocument.Open(presentationFile, false);
        presentationDocument.ExportToQuartoMarkdown(
            "/Users/nunkesser/repos/work/ppt/PM-Folien",
            fileName);
    }
}