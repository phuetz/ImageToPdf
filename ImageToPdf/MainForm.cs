using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using Markdig;

namespace ImageToPdf;

public class MainForm : Form
{
    private ListBox listBoxFiles;
    private Button btnAddFiles;
    private Button btnRemoveSelected;
    private Button btnMoveUp;
    private Button btnMoveDown;
    private Button btnClear;
    private Button btnConvert;
    private Label lblInfo;
    private ProgressBar progressBar;

    private List<string> filePaths = new();

    private static readonly string[] ImageExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif" };
    private static readonly string[] PdfExtensions = { ".pdf" };
    private static readonly string[] MarkdownExtensions = { ".md", ".markdown" };

    public MainForm()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        this.Text = "PDF Merger - Images, PDF & Markdown";
        this.Size = new Size(650, 500);
        this.MinimumSize = new Size(550, 400);
        this.StartPosition = FormStartPosition.CenterScreen;

        // Label info
        lblInfo = new Label
        {
            Text = "Ajoutez des images, PDF ou fichiers Markdown à fusionner:",
            Location = new Point(12, 12),
            Size = new Size(400, 20)
        };

        // ListBox pour les fichiers
        listBoxFiles = new ListBox
        {
            Location = new Point(12, 35),
            Size = new Size(440, 350),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            SelectionMode = SelectionMode.MultiExtended,
            HorizontalScrollbar = true
        };

        // Boutons
        int btnX = 460;
        int btnWidth = 160;

        btnAddFiles = new Button
        {
            Text = "Ajouter fichiers...",
            Location = new Point(btnX, 35),
            Size = new Size(btnWidth, 30),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        btnAddFiles.Click += BtnAddFiles_Click;

        btnRemoveSelected = new Button
        {
            Text = "Supprimer sélection",
            Location = new Point(btnX, 70),
            Size = new Size(btnWidth, 30),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        btnRemoveSelected.Click += BtnRemoveSelected_Click;

        btnMoveUp = new Button
        {
            Text = "Monter ↑",
            Location = new Point(btnX, 115),
            Size = new Size(btnWidth, 30),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        btnMoveUp.Click += BtnMoveUp_Click;

        btnMoveDown = new Button
        {
            Text = "Descendre ↓",
            Location = new Point(btnX, 150),
            Size = new Size(btnWidth, 30),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        btnMoveDown.Click += BtnMoveDown_Click;

        btnClear = new Button
        {
            Text = "Tout effacer",
            Location = new Point(btnX, 195),
            Size = new Size(btnWidth, 30),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        btnClear.Click += BtnClear_Click;

        btnConvert = new Button
        {
            Text = "Créer le PDF",
            Location = new Point(btnX, 280),
            Size = new Size(btnWidth, 45),
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        btnConvert.Click += BtnConvert_Click;

        // ProgressBar
        progressBar = new ProgressBar
        {
            Location = new Point(12, 400),
            Size = new Size(606, 25),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            Visible = false
        };

        // Ajouter les contrôles
        this.Controls.AddRange(new Control[]
        {
            lblInfo, listBoxFiles, btnAddFiles, btnRemoveSelected,
            btnMoveUp, btnMoveDown, btnClear, btnConvert, progressBar
        });

        // Support du drag & drop
        this.AllowDrop = true;
        this.DragEnter += MainForm_DragEnter;
        this.DragDrop += MainForm_DragDrop;
    }

    private void MainForm_DragEnter(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
        {
            e.Effect = DragDropEffects.Copy;
        }
    }

    private void MainForm_DragDrop(object? sender, DragEventArgs e)
    {
        var files = e.Data?.GetData(DataFormats.FileDrop) as string[];
        if (files != null)
        {
            AddFiles(files);
        }
    }

    private void BtnAddFiles_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Sélectionner des fichiers",
            Filter = "Tous les fichiers supportés|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif;*.pdf;*.md;*.markdown|" +
                     "Images|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif|" +
                     "PDF|*.pdf|" +
                     "Markdown|*.md;*.markdown|" +
                     "Tous les fichiers|*.*",
            Multiselect = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            AddFiles(dialog.FileNames);
        }
    }

    private void AddFiles(string[] files)
    {
        var allValidExtensions = ImageExtensions.Concat(PdfExtensions).Concat(MarkdownExtensions).ToArray();

        foreach (var file in files)
        {
            var ext = Path.GetExtension(file).ToLowerInvariant();
            if (allValidExtensions.Contains(ext) && !filePaths.Contains(file))
            {
                filePaths.Add(file);
                var fileType = GetFileType(ext);
                listBoxFiles.Items.Add($"[{fileType}] {Path.GetFileName(file)}");
            }
        }

        UpdateTitle();
    }

    private static string GetFileType(string extension)
    {
        if (ImageExtensions.Contains(extension)) return "IMG";
        if (PdfExtensions.Contains(extension)) return "PDF";
        if (MarkdownExtensions.Contains(extension)) return "MD";
        return "???";
    }

    private void BtnRemoveSelected_Click(object? sender, EventArgs e)
    {
        var indices = listBoxFiles.SelectedIndices.Cast<int>().OrderByDescending(i => i).ToList();
        foreach (var index in indices)
        {
            filePaths.RemoveAt(index);
            listBoxFiles.Items.RemoveAt(index);
        }
        UpdateTitle();
    }

    private void BtnMoveUp_Click(object? sender, EventArgs e)
    {
        if (listBoxFiles.SelectedIndex > 0)
        {
            int index = listBoxFiles.SelectedIndex;
            SwapItems(index, index - 1);
            listBoxFiles.SelectedIndex = index - 1;
        }
    }

    private void BtnMoveDown_Click(object? sender, EventArgs e)
    {
        if (listBoxFiles.SelectedIndex >= 0 && listBoxFiles.SelectedIndex < listBoxFiles.Items.Count - 1)
        {
            int index = listBoxFiles.SelectedIndex;
            SwapItems(index, index + 1);
            listBoxFiles.SelectedIndex = index + 1;
        }
    }

    private void SwapItems(int index1, int index2)
    {
        (filePaths[index1], filePaths[index2]) = (filePaths[index2], filePaths[index1]);

        var temp = listBoxFiles.Items[index1];
        listBoxFiles.Items[index1] = listBoxFiles.Items[index2];
        listBoxFiles.Items[index2] = temp;
    }

    private void BtnClear_Click(object? sender, EventArgs e)
    {
        filePaths.Clear();
        listBoxFiles.Items.Clear();
        UpdateTitle();
    }

    private void UpdateTitle()
    {
        this.Text = $"PDF Merger - {filePaths.Count} fichier(s)";
    }

    private async void BtnConvert_Click(object? sender, EventArgs e)
    {
        if (filePaths.Count == 0)
        {
            MessageBox.Show("Veuillez ajouter au moins un fichier.", "Aucun fichier",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var saveDialog = new SaveFileDialog
        {
            Title = "Enregistrer le PDF",
            Filter = "PDF|*.pdf",
            DefaultExt = "pdf",
            FileName = "merged.pdf"
        };

        if (saveDialog.ShowDialog() != DialogResult.OK)
            return;

        progressBar.Visible = true;
        progressBar.Value = 0;
        progressBar.Maximum = filePaths.Count;
        SetButtonsEnabled(false);

        try
        {
            await Task.Run(() => CreatePdf(saveDialog.FileName));

            MessageBox.Show($"PDF créé avec succès!\n{saveDialog.FileName}", "Succès",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Erreur lors de la création du PDF:\n{ex.Message}", "Erreur",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            progressBar.Visible = false;
            SetButtonsEnabled(true);
        }
    }

    private void SetButtonsEnabled(bool enabled)
    {
        btnAddFiles.Enabled = enabled;
        btnRemoveSelected.Enabled = enabled;
        btnMoveUp.Enabled = enabled;
        btnMoveDown.Enabled = enabled;
        btnClear.Enabled = enabled;
        btnConvert.Enabled = enabled;
    }

    private void CreatePdf(string outputPath)
    {
        using var document = new PdfDocument();
        document.Info.Title = "Document fusionné";
        document.Info.Creator = "PDF Merger";

        for (int i = 0; i < filePaths.Count; i++)
        {
            var filePath = filePaths[i];
            var ext = Path.GetExtension(filePath).ToLowerInvariant();

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
            }

            this.Invoke(() =>
            {
                progressBar.Value = i + 1;
            });
        }

        document.Save(outputPath);
    }

    private static void AddImageToPdf(PdfDocument document, string imagePath)
    {
        using var stream = File.OpenRead(imagePath);
        var image = XImage.FromStream(() => stream);

        var page = document.AddPage();
        page.Width = XUnit.FromPoint(image.PointWidth);
        page.Height = XUnit.FromPoint(image.PointHeight);

        using var gfx = XGraphics.FromPdfPage(page);
        gfx.DrawImage(image, 0, 0, page.Width, page.Height);
    }

    private static void AddPdfToPdf(PdfDocument document, string pdfPath)
    {
        using var inputDocument = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);

        for (int i = 0; i < inputDocument.PageCount; i++)
        {
            var page = inputDocument.Pages[i];
            document.AddPage(page);
        }
    }

    private static void AddMarkdownToPdf(PdfDocument document, string markdownPath)
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
