namespace Italbytz.OpenXml.Tests;

[TestClass]
public sealed class PresentationConverterTest
{
    [TestMethod]
    public void TestMethod4()
    {
        PresentationConverter.ConvertToQuartoMarkdown(
            "/Users/nunkesser/repos/work/ppt/CES.pptx",
            "/Users/nunkesser/repos/work/md/quarto/gdi/CES.qmd",
            "/Users/nunkesser/repos/work/images/gdi", "XYZ",
            "Prof. Dr. Robin Nunkesser");
    }
}