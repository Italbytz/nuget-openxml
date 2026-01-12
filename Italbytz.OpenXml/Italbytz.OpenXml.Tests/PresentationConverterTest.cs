namespace Italbytz.OpenXml.Tests;

[TestClass]
public sealed class PresentationConverterTest
{
    [TestMethod]
    public void TestMethod4()
    {
        PresentationConverter.ConvertToQuartoMarkdown(
            "/Users/nunkesser/repos/work/ppt/PM-Folien/LE01.pptx",
            "/Users/nunkesser/repos/work/md/quarto/itp/Einfuehrung.qmd",
            "/Users/nunkesser/repos/work/images/itp", "Einführung",
            "Prof. Dr. Robin Nunkesser");
    }
}