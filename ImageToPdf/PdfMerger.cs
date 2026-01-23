using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf.IO;
using Markdig;
using PdfSharpDocument = PdfSharpCore.Pdf.PdfDocument;

namespace ImageToPdf;

public static class PdfMerger
{
    private static readonly string[] ImageExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif" };
    private static readonly string[] PdfExtensions = { ".pdf" };
    private static readonly string[] MarkdownExtensions = { ".md", ".markdown" };

    public static bool IsValidExtension(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ImageExtensions.Contains(ext) || PdfExtensions.Contains(ext) || MarkdownExtensions.Contains(ext);
    }

    public static void CreatePdf(List<string> filePaths, string outputPath, Action<int, int, string>? progress = null)
    {
        using var document = new PdfSharpDocument();
        document.Info.Title = "Document fusionné";
        document.Info.Creator = "PDF Merger";

        for (int i = 0; i < filePaths.Count; i++)
        {
            var filePath = filePaths[i];
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            var fileName = Path.GetFileName(filePath);

            progress?.Invoke(i + 1, filePaths.Count, fileName);

            try
            {
                if (ImageExtensions.Contains(ext))
                {
                    AddImageToPdf(document, filePath);
                }
                else if (PdfExtensions.Contains(ext))
                {
                    AddPdfToPdf(document, filePath);
                }
                else if (MarkdownExtensions.Contains(ext))
                {
                    AddMarkdownToPdf(document, filePath);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur avec {filePath}: {ex.Message}");
                throw new Exception($"Erreur lors du traitement de '{fileName}': {ex.Message}", ex);
            }
        }

        document.Save(outputPath);
    }

    private static void AddImageToPdf(PdfSharpDocument document, string imagePath)
    {
        using var stream = File.OpenRead(imagePath);
        var image = XImage.FromStream(() => stream);

        var page = document.AddPage();
        page.Width = XUnit.FromPoint(image.PointWidth);
        page.Height = XUnit.FromPoint(image.PointHeight);

        using var gfx = XGraphics.FromPdfPage(page);
        gfx.DrawImage(image, 0, 0, page.Width, page.Height);
    }

    private static void AddPdfToPdf(PdfSharpDocument document, string pdfPath)
    {
        using var inputDocument = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);

        for (int i = 0; i < inputDocument.PageCount; i++)
        {
            var page = inputDocument.Pages[i];
            document.AddPage(page);
        }
    }

    private static void AddMarkdownToPdf(PdfSharpDocument document, string markdownPath)
    {
        var markdownContent = File.ReadAllText(markdownPath);
        var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
        var plainText = Markdown.ToPlainText(markdownContent, pipeline);

        var page = document.AddPage();
        page.Size = PdfSharpCore.PageSize.A4;

        using var gfx = XGraphics.FromPdfPage(page);
        var font = new XFont("Arial", 11);
        var titleFont = new XFont("Arial", 16, XFontStyle.Bold);

        double margin = 50;
        double y = margin;
        double lineHeight = 16;
        double maxWidth = page.Width - (2 * margin);

        // Titre du fichier
        var fileName = Path.GetFileNameWithoutExtension(markdownPath);
        gfx.DrawString(fileName, titleFont, XBrushes.Black, margin, y);
        y += 30;

        // Contenu
        var lines = plainText.Split('\n');
        foreach (var line in lines)
        {
            if (y > page.Height - margin)
            {
                page = document.AddPage();
                page.Size = PdfSharpCore.PageSize.A4;
                gfx.Dispose();
                var newGfx = XGraphics.FromPdfPage(page);
                y = margin;
                DrawTextLine(newGfx, line.Trim(), font, margin, ref y, lineHeight, maxWidth, page.Height - margin);
                newGfx.Dispose();
            }
            else
            {
                DrawTextLine(gfx, line.Trim(), font, margin, ref y, lineHeight, maxWidth, page.Height - margin);
            }
        }
    }

    private static void DrawTextLine(XGraphics gfx, string text, XFont font, double x, ref double y, double lineHeight, double maxWidth, double maxY)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            y += lineHeight / 2;
            return;
        }

        var words = text.Split(' ');
        var currentLine = "";

        foreach (var word in words)
        {
            var testLine = string.IsNullOrEmpty(currentLine) ? word : currentLine + " " + word;
            var size = gfx.MeasureString(testLine, font);

            if (size.Width > maxWidth && !string.IsNullOrEmpty(currentLine))
            {
                gfx.DrawString(currentLine, font, XBrushes.Black, x, y);
                y += lineHeight;
                currentLine = word;

                if (y > maxY)
                    return;
            }
            else
            {
                currentLine = testLine;
            }
        }

        if (!string.IsNullOrEmpty(currentLine))
        {
            gfx.DrawString(currentLine, font, XBrushes.Black, x, y);
            y += lineHeight;
        }
    }
}
