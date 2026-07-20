using System.Text;
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

    // Vrai uniquement pour les formats image (pas PDF ni Markdown) — utilisé par la
    // conversion par lot qui, par définition, ne regroupe que des images par dossier.
    public static bool IsImageFile(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ImageExtensions.Contains(ext);
    }

    public static void CreatePdf(List<string> filePaths, string outputPath, Action<int, int, string>? progress = null, CancellationToken cancellationToken = default)
    {
        using var document = new PdfSharpDocument();
        document.Info.Title = "Document fusionné";
        document.Info.Creator = "PDF Merger";

        for (int i = 0; i < filePaths.Count; i++)
        {
            // Annulation coopérative : levée hors du try/catch par fichier pour
            // remonter telle quelle (l'appelant distingue annulation et erreur).
            cancellationToken.ThrowIfCancellationRequested();

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
        var extension = Path.GetExtension(imagePath).ToLowerInvariant();
        if (extension is ".tif" or ".tiff")
        {
            AddTiffToPdf(document, imagePath);
            return;
        }

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

    private static void AddTiffToPdf(PdfSharpDocument document, string tiffPath)
    {
        // GDI+ décode les TIFF multi-pages (sortie typique des scanners) ; ImageSharp,
        // le décodeur interne de PdfSharpCore, ne les gère pas. Une page PDF par frame,
        // chaque frame ré-encodée en PNG en mémoire — sans la boucle, seules la
        // première page d'un scan de N pages arrivait dans le PDF.
        using var img = System.Drawing.Image.FromFile(tiffPath);
        var pageDimension = new System.Drawing.Imaging.FrameDimension(img.FrameDimensionsList[0]);
        int frameCount;
        try { frameCount = img.GetFrameCount(pageDimension); }
        catch { frameCount = 1; }

        for (int frame = 0; frame < frameCount; frame++)
        {
            if (frameCount > 1)
                img.SelectActiveFrame(pageDimension, frame);

            using var ms = new MemoryStream();
            img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            ms.Position = 0;

            using var image = XImage.FromStream(() => ms);
            var page = document.AddPage();
            page.Width = XUnit.FromPoint(image.PointWidth);
            page.Height = XUnit.FromPoint(image.PointHeight);

            using var gfx = XGraphics.FromPdfPage(page);
            gfx.DrawImage(image, 0, 0, page.Width, page.Height);
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
        var markdownContent = ReadAllTextSmart(markdownPath);
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

    // Lit un fichier texte en détectant l'encodage : BOM d'abord, sinon UTF-8 strict,
    // avec repli Windows-1252 (« ANSI » français) — sans quoi les accents des fichiers
    // issus d'éditeurs anciens deviennent « � » jusque dans le PDF final.
    public static string ReadAllTextSmart(string path)
    {
        var bytes = File.ReadAllBytes(path);

        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);

        try
        {
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true).GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            return Encoding.GetEncoding(1252).GetString(bytes);
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
