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
        XImage? image = null;
        Stream? keepAlive = null; // le flux doit rester ouvert jusqu'au DrawImage
        try
        {
            try
            {
                // Chemin natif (rapide, conserve la compression d'origine).
                var fileStream = File.OpenRead(imagePath);
                keepAlive = fileStream;
                image = XImage.FromStream(() => fileStream);
                _ = image.PixelWidth; // force le décodage pour détecter un format non géré (ex: TIFF)
            }
            catch
            {
                // Repli GDI+ : le décodeur interne (ImageSharp) ne gère pas tous les
                // formats annoncés — TIFF notamment. System.Drawing les lit, on
                // ré-encode en PNG en mémoire pour PdfSharpCore.
                image?.Dispose();
                keepAlive?.Dispose();

                var ms = new MemoryStream();
                using (var bmp = new System.Drawing.Bitmap(imagePath))
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;
                keepAlive = ms;
                image = XImage.FromStream(() => ms);
            }

            var page = document.AddPage();
            page.Width = XUnit.FromPoint(image.PointWidth);
            page.Height = XUnit.FromPoint(image.PointHeight);

            using var gfx = XGraphics.FromPdfPage(page);
            gfx.DrawImage(image, 0, 0, page.Width, page.Height);
        }
        finally
        {
            image?.Dispose();
            keepAlive?.Dispose();
        }
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

        var gfx = XGraphics.FromPdfPage(page);
        try
        {
            var font = new XFont("Arial", 11);
            var titleFont = new XFont("Arial", 16, XFontStyle.Bold);

            double margin = 50;
            double y = margin;
            double lineHeight = 16;
            double maxWidth = page.Width - (2 * margin);
            double bottom = page.Height - margin;

            // Titre du fichier
            var fileName = Path.GetFileNameWithoutExtension(markdownPath);
            gfx.DrawString(fileName, titleFont, XBrushes.Black, margin, y);
            y += 30;

            // 1) On calcule d'abord TOUTES les lignes d'affichage (retour à la ligne
            //    mot à mot). La mesure est indépendante de la page (toutes en A4),
            //    donc on peut la faire avec le gfx courant avant toute pagination.
            //    null = ligne vide (demi-interligne).
            var displayLines = new List<string?>();
            foreach (var rawLine in plainText.Split('\n'))
            {
                var line = rawLine.Trim();
                if (string.IsNullOrWhiteSpace(line))
                {
                    displayLines.Add(null);
                    continue;
                }
                foreach (var wrapped in WrapLine(gfx, line, font, maxWidth))
                    displayLines.Add(wrapped);
            }

            // 2) On dessine en paginant proprement (une nouvelle page dès qu'on
            //    dépasse le bas), sans jamais perdre de contenu.
            foreach (var dl in displayLines)
            {
                if (dl == null)
                {
                    y += lineHeight / 2;
                    continue;
                }

                if (y > bottom)
                {
                    gfx.Dispose();
                    page = document.AddPage();
                    page.Size = PdfSharpCore.PageSize.A4;
                    gfx = XGraphics.FromPdfPage(page);
                    y = margin;
                }

                gfx.DrawString(dl, font, XBrushes.Black, margin, y);
                y += lineHeight;
            }
        }
        finally
        {
            gfx.Dispose();
        }
    }

    // Découpe une ligne en lignes d'affichage tenant dans maxWidth (retour à la ligne mot à mot).
    private static IEnumerable<string> WrapLine(XGraphics gfx, string text, XFont font, double maxWidth)
    {
        var words = text.Split(' ');
        var currentLine = "";

        foreach (var word in words)
        {
            var testLine = string.IsNullOrEmpty(currentLine) ? word : currentLine + " " + word;

            if (gfx.MeasureString(testLine, font).Width > maxWidth && !string.IsNullOrEmpty(currentLine))
            {
                yield return currentLine;
                currentLine = word;
            }
            else
            {
                currentLine = testLine;
            }
        }

        if (!string.IsNullOrEmpty(currentLine))
            yield return currentLine;
    }
}
