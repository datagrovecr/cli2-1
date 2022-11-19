namespace convert_test;
using OfficeConvert;
using Markdig;
using System;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.Text;
using System.IO.Compression;
using System.Linq.Expressions;
using DocumentFormat.OpenXml.EMMA;
using System.Diagnostics;

class CustomException : Exception
{
    public CustomException(string message)
    {
    }
}

[TestClass]
public class UnitTest1
{
    private TestContext testContextInstance;

    /// <summary>
    /// Gets or sets the test context which provides
    /// information about and functionality for the current test run.
    /// </summary>
    public TestContext TestContext
    {
        get { return testContextInstance; }
        set { testContextInstance = value; }
    }
    [TestMethod]
    public async Task ConversionTests1()
    {
        var d = Directory.GetCurrentDirectory()!+ "/../../../";
        await OfficeConvert.Convert(d+"tests/one", d+"TestResults/tests", d+"TestResults/round", true);
    }


    [TestMethod]
    public async Task ConversionTests()
    {
        var d = Directory.GetCurrentDirectory()!+ "/../../../";
        await OfficeConvert.Convert(d+"tests", d+"TestResults/tests", d+"TestResults/round", true);
    }

    [TestMethod]
    public static void createWordprocessingDocument(string filepath)
    {
        // Create a document by supplying the filepath. 
        using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
        {
            // Add a main document part. 
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
        }

    }
}