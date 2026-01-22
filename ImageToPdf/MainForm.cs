using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf.IO;
using Markdig;
using System.Drawing.Drawing2D;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using PdfSharpDocument = PdfSharpCore.Pdf.PdfDocument;
using WordDocument = DocumentFormat.OpenXml.Wordprocessing.Document;
using WordBody = DocumentFormat.OpenXml.Wordprocessing.Body;
using WordParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WordRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace ImageToPdf;

public class MainForm : Form
{
    private ListView listViewFiles = null!;
    private ImageList imageListIcons = null!;
    private Button btnAddFiles = null!;
    private Button btnRemoveSelected = null!;
    private Button btnMoveUp = null!;
    private Button btnMoveDown = null!;
    private Button btnClear = null!;
    private Button btnConvert = null!;
    private Button btnPreviewResult = null!;
    private Button btnTogglePreview = null!;
    private Label lblInfo = null!;
    private ProgressBar progressBar = null!;
    private MenuStrip menuStrip = null!;

    // Panels pour la disposition
    private Panel mainPanel = null!;
    private Panel buttonPanel = null!;
    private Panel contentPanel = null!;

    // Panneau de prévisualisation
    private Panel previewPanel = null!;
    private PictureBox pictureBoxPreview = null!;
    private RichTextBox textBoxPreview = null!;
    private Label lblPreviewInfo = null!;
    private Label lblNoPreview = null!;
    private Splitter splitter = null!;

    private List<string> filePaths = new();
    private bool previewVisible = false;

    private static readonly string[] ImageExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif" };
    private static readonly string[] PdfExtensions = { ".pdf" };
    private static readonly string[] MarkdownExtensions = { ".md", ".markdown" };

    private const int ICON_IMAGE = 0;
    private const int ICON_PDF = 1;
    private const int ICON_MARKDOWN = 2;

    public MainForm()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        this.Text = "PDF Merger Standard";
        this.Size = new Size(650, 550);
        this.MinimumSize = new Size(550, 450);
        this.StartPosition = FormStartPosition.CenterScreen;

        CreateMenuStrip();
        CreateImageList();
        CreatePreviewPanel();
        CreateMainPanel();

        // Splitter entre main et preview
        splitter = new Splitter
        {
            Dock = DockStyle.Right,
            Width = 4,
            BackColor = Color.FromArgb(200, 200, 200),
            Visible = false
        };

        // Ordre important : d'abord les éléments Dock.Right, puis Dock.Fill
        this.Controls.Add(previewPanel);
        this.Controls.Add(splitter);
        this.Controls.Add(mainPanel);
        this.Controls.Add(menuStrip);
        this.MainMenuStrip = menuStrip;

        // Support du drag & drop
        this.AllowDrop = true;
        this.DragEnter += MainForm_DragEnter;
        this.DragDrop += MainForm_DragDrop;
    }

    private void CreateMenuStrip()
    {
        menuStrip = new MenuStrip();

        // Menu Outils
        var toolsMenu = new ToolStripMenuItem("Outils");

        var pdfToWordItem = new ToolStripMenuItem("PDF → Word", null, PdfToWord_Click)
        {
            ShortcutKeys = Keys.Control | Keys.W
        };

        var openPdfSamItem = new ToolStripMenuItem("Ouvrir PDFsam", null, OpenPdfSam_Click)
        {
            ShortcutKeys = Keys.Control | Keys.P
        };

        toolsMenu.DropDownItems.Add(pdfToWordItem);
        toolsMenu.DropDownItems.Add(new ToolStripSeparator());
        toolsMenu.DropDownItems.Add(openPdfSamItem);

        menuStrip.Items.Add(toolsMenu);
    }

    private void CreateMainPanel()
    {
        mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(10)
        };

        // Label info en haut
        lblInfo = new Label
        {
            Text = "Ajoutez des images, PDF ou fichiers Markdown à fusionner:",
            Dock = DockStyle.Top,
            Height = 25,
            Padding = new Padding(0, 5, 0, 0)
        };

        // Panneau des boutons à droite
        buttonPanel = new Panel
        {
            Dock = DockStyle.Right,
            Width = 140,
            Padding = new Padding(10, 0, 0, 0)
        };

        int btnWidth = 125;
        int btnY = 0;
        int btnSpacing = 32;

        btnAddFiles = new Button
        {
            Text = "Ajouter...",
            Location = new Point(10, btnY),
            Size = new Size(btnWidth, 28)
        };
        btnAddFiles.Click += BtnAddFiles_Click;
        btnY += btnSpacing;

        btnRemoveSelected = new Button
        {
            Text = "Supprimer",
            Location = new Point(10, btnY),
            Size = new Size(btnWidth, 28)
        };
        btnRemoveSelected.Click += BtnRemoveSelected_Click;
        btnY += btnSpacing + 10;

        btnMoveUp = new Button
        {
            Text = "▲ Monter",
            Location = new Point(10, btnY),
            Size = new Size(btnWidth, 28)
        };
        btnMoveUp.Click += BtnMoveUp_Click;
        btnY += btnSpacing;

        btnMoveDown = new Button
        {
            Text = "▼ Descendre",
            Location = new Point(10, btnY),
            Size = new Size(btnWidth, 28)
        };
        btnMoveDown.Click += BtnMoveDown_Click;
        btnY += btnSpacing + 10;

        btnClear = new Button
        {
            Text = "Tout effacer",
            Location = new Point(10, btnY),
            Size = new Size(btnWidth, 28)
        };
        btnClear.Click += BtnClear_Click;
        btnY += btnSpacing + 20;

        btnTogglePreview = new Button
        {
            Text = "Aperçu ▶",
            Location = new Point(10, btnY),
            Size = new Size(btnWidth, 28),
            BackColor = Color.FromArgb(240, 240, 240)
        };
        btnTogglePreview.Click += BtnTogglePreview_Click;
        btnY += btnSpacing + 20;

        btnPreviewResult = new Button
        {
            Text = "Voir résultat",
            Location = new Point(10, btnY),
            Size = new Size(btnWidth, 28),
            BackColor = Color.FromArgb(76, 175, 80),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        btnPreviewResult.Click += BtnPreviewResult_Click;
        btnY += btnSpacing + 10;

        btnConvert = new Button
        {
            Text = "Créer le PDF",
            Location = new Point(10, btnY),
            Size = new Size(btnWidth, 42),
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font(this.Font.FontFamily, 9, FontStyle.Bold)
        };
        btnConvert.Click += BtnConvert_Click;

        buttonPanel.Controls.AddRange(new Control[]
        {
            btnAddFiles, btnRemoveSelected, btnMoveUp, btnMoveDown,
            btnClear, btnTogglePreview, btnPreviewResult, btnConvert
        });

        // Panneau du contenu (liste + progress bar)
        contentPanel = new Panel
        {
            Dock = DockStyle.Fill
        };

        // ListView
        listViewFiles = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            SmallImageList = imageListIcons,
            MultiSelect = true,
            HeaderStyle = ColumnHeaderStyle.Nonclickable,
            BorderStyle = BorderStyle.FixedSingle
        };
        listViewFiles.Columns.Add("Fichier", 280);
        listViewFiles.Columns.Add("Type", 70);
        listViewFiles.SelectedIndexChanged += ListViewFiles_SelectedIndexChanged;

        // ProgressBar
        progressBar = new ProgressBar
        {
            Dock = DockStyle.Bottom,
            Height = 23,
            Visible = false
        };

        contentPanel.Controls.Add(listViewFiles);
        contentPanel.Controls.Add(progressBar);

        // Assemblage du panneau principal
        mainPanel.Controls.Add(contentPanel);
        mainPanel.Controls.Add(buttonPanel);
        mainPanel.Controls.Add(lblInfo);
    }

    private void CreateImageList()
    {
        imageListIcons = new ImageList
        {
            ImageSize = new Size(20, 20),
            ColorDepth = ColorDepth.Depth32Bit
        };

        imageListIcons.Images.Add("image", CreateImageIcon());
        imageListIcons.Images.Add("pdf", CreatePdfIcon());
        imageListIcons.Images.Add("markdown", CreateMarkdownIcon());
    }

    private static Bitmap CreateImageIcon()
    {
        var bmp = new Bitmap(20, 20);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);

        // Cadre de l'image
        using var framePen = new Pen(Color.FromArgb(100, 100, 100), 1);
        using var frameBrush = new SolidBrush(Color.FromArgb(240, 240, 240));
        g.FillRectangle(frameBrush, 1, 1, 18, 18);
        g.DrawRectangle(framePen, 1, 1, 17, 17);

        // Ciel bleu
        using var skyBrush = new SolidBrush(Color.FromArgb(135, 206, 250));
        g.FillRectangle(skyBrush, 2, 2, 16, 10);

        // Soleil
        using var sunBrush = new SolidBrush(Color.FromArgb(255, 200, 50));
        g.FillEllipse(sunBrush, 12, 3, 5, 5);

        // Montagne verte
        var mountain = new Point[] { new(2, 17), new(8, 8), new(14, 17) };
        using var mountainBrush = new SolidBrush(Color.FromArgb(76, 175, 80));
        g.FillPolygon(mountainBrush, mountain);

        // Petite montagne
        var smallMountain = new Point[] { new(10, 17), new(15, 10), new(18, 17) };
        using var smallMountainBrush = new SolidBrush(Color.FromArgb(56, 142, 60));
        g.FillPolygon(smallMountainBrush, smallMountain);

        return bmp;
    }

    private static Bitmap CreatePdfIcon()
    {
        var bmp = new Bitmap(20, 20);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);

        // Document blanc avec coin plié
        var docPoints = new Point[] { new(3, 1), new(14, 1), new(17, 4), new(17, 18), new(3, 18) };
        using var docBrush = new SolidBrush(Color.White);
        using var docPen = new Pen(Color.FromArgb(200, 50, 50), 1);
        g.FillPolygon(docBrush, docPoints);
        g.DrawPolygon(docPen, docPoints);

        // Coin plié
        var foldPoints = new Point[] { new(14, 1), new(14, 4), new(17, 4) };
        using var foldBrush = new SolidBrush(Color.FromArgb(230, 230, 230));
        g.FillPolygon(foldBrush, foldPoints);
        g.DrawPolygon(docPen, foldPoints);

        // Bandeau rouge PDF
        using var pdfBrush = new SolidBrush(Color.FromArgb(220, 50, 50));
        g.FillRectangle(pdfBrush, 3, 10, 14, 7);

        // Texte PDF
        using var font = new Font("Arial", 6, FontStyle.Bold);
        using var textBrush = new SolidBrush(Color.White);
        g.DrawString("PDF", font, textBrush, 4, 10);

        return bmp;
    }

    private static Bitmap CreateMarkdownIcon()
    {
        var bmp = new Bitmap(20, 20);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);

        // Document
        using var docBrush = new SolidBrush(Color.White);
        using var docPen = new Pen(Color.FromArgb(50, 50, 50), 1);
        g.FillRectangle(docBrush, 2, 1, 16, 18);
        g.DrawRectangle(docPen, 2, 1, 15, 17);

        // Fond bleu pour le symbole
        using var mdBgBrush = new SolidBrush(Color.FromArgb(33, 150, 243));
        g.FillRectangle(mdBgBrush, 4, 5, 12, 10);

        // Symbole M avec flèche (Markdown)
        using var mdPen = new Pen(Color.White, 1.5f);

        // M
        g.DrawLine(mdPen, 6, 12, 6, 7);
        g.DrawLine(mdPen, 6, 7, 8, 10);
        g.DrawLine(mdPen, 8, 10, 10, 7);
        g.DrawLine(mdPen, 10, 7, 10, 12);

        // Flèche vers le bas
        g.DrawLine(mdPen, 13, 7, 13, 12);
        g.DrawLine(mdPen, 11, 10, 13, 12);
        g.DrawLine(mdPen, 15, 10, 13, 12);

        return bmp;
    }

    private void CreatePreviewPanel()
    {
        previewPanel = new Panel
        {
            Dock = DockStyle.Right,
            Width = 320,
            Visible = false,
            BackColor = Color.FromArgb(248, 248, 248),
            Padding = new Padding(10)
        };

        lblPreviewInfo = new Label
        {
            Text = "Aperçu",
            Dock = DockStyle.Top,
            Height = 30,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50),
            Padding = new Padding(0, 5, 0, 0)
        };

        var previewContentPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 5, 0, 0)
        };

        pictureBoxPreview = new PictureBox
        {
            Dock = DockStyle.Fill,
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Visible = false
        };

        textBoxPreview = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Consolas", 9),
            Visible = false
        };

        lblNoPreview = new Label
        {
            Dock = DockStyle.Fill,
            Text = "Sélectionnez un fichier\npour voir l'aperçu",
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 10),
            BackColor = Color.White
        };

        previewContentPanel.Controls.Add(pictureBoxPreview);
        previewContentPanel.Controls.Add(textBoxPreview);
        previewContentPanel.Controls.Add(lblNoPreview);

        previewPanel.Controls.Add(previewContentPanel);
        previewPanel.Controls.Add(lblPreviewInfo);
    }

    private void BtnTogglePreview_Click(object? sender, EventArgs e)
    {
        previewVisible = !previewVisible;
        previewPanel.Visible = previewVisible;
        splitter.Visible = previewVisible;

        if (previewVisible)
        {
            btnTogglePreview.Text = "◀ Masquer";
            if (this.Width < 900)
            {
                this.Width = 950;
            }
            UpdatePreview();
        }
        else
        {
            btnTogglePreview.Text = "Aperçu ▶";
            if (this.Width > 700)
            {
                this.Width = 650;
            }
        }
    }

    private void ListViewFiles_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (previewVisible)
        {
            UpdatePreview();
        }
    }

    private void UpdatePreview()
    {
        pictureBoxPreview.Visible = false;
        textBoxPreview.Visible = false;
        lblNoPreview.Visible = true;
        pictureBoxPreview.Image?.Dispose();
        pictureBoxPreview.Image = null;

        if (listViewFiles.SelectedIndices.Count == 0)
        {
            lblPreviewInfo.Text = "Aperçu";
            lblNoPreview.Text = "Sélectionnez un fichier\npour voir l'aperçu";
            return;
        }

        var index = listViewFiles.SelectedIndices[0];
        var filePath = filePaths[index];
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var fileName = Path.GetFileName(filePath);

        lblPreviewInfo.Text = fileName.Length > 35 ? fileName[..32] + "..." : fileName;

        try
        {
            if (ImageExtensions.Contains(ext))
            {
                ShowImagePreview(filePath);
            }
            else if (PdfExtensions.Contains(ext))
            {
                ShowPdfPreview(filePath);
            }
            else if (MarkdownExtensions.Contains(ext))
            {
                ShowMarkdownPreview(filePath);
            }
        }
        catch (Exception ex)
        {
            lblNoPreview.Text = $"Erreur:\n{ex.Message}";
            lblNoPreview.Visible = true;
        }
    }

    private void ShowImagePreview(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        var img = Image.FromStream(stream);
        pictureBoxPreview.Image = new Bitmap(img);
        img.Dispose();

        pictureBoxPreview.Visible = true;
        lblNoPreview.Visible = false;
    }

    private void ShowPdfPreview(string filePath)
    {
        using var doc = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
        var pageCount = doc.PageCount;
        var firstPage = doc.Pages[0];
        var width = firstPage.Width.Point;
        var height = firstPage.Height.Point;

        textBoxPreview.Clear();
        textBoxPreview.SelectionFont = new Font("Segoe UI", 12, FontStyle.Bold);
        textBoxPreview.AppendText("Document PDF\n\n");
        textBoxPreview.SelectionFont = new Font("Segoe UI", 10);
        textBoxPreview.AppendText($"  Pages: {pageCount}\n\n");
        textBoxPreview.AppendText($"  Dimensions: {width:F0} x {height:F0} pt\n\n");

        var fileInfo = new FileInfo(filePath);
        var sizeKb = fileInfo.Length / 1024.0;
        var sizeStr = sizeKb > 1024 ? $"{sizeKb / 1024:F1} Mo" : $"{sizeKb:F0} Ko";
        textBoxPreview.AppendText($"  Taille: {sizeStr}\n");

        textBoxPreview.Visible = true;
        lblNoPreview.Visible = false;
    }

    private void ShowMarkdownPreview(string filePath)
    {
        var content = File.ReadAllText(filePath);

        textBoxPreview.Clear();
        textBoxPreview.Text = content;
        textBoxPreview.Visible = true;
        lblNoPreview.Visible = false;
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
                var (iconIndex, typeName) = GetFileTypeInfo(ext);

                var item = new ListViewItem(Path.GetFileName(file), iconIndex);
                item.SubItems.Add(typeName);
                listViewFiles.Items.Add(item);
            }
        }

        UpdateTitle();
    }

    private static (int iconIndex, string typeName) GetFileTypeInfo(string extension)
    {
        if (ImageExtensions.Contains(extension)) return (ICON_IMAGE, "Image");
        if (PdfExtensions.Contains(extension)) return (ICON_PDF, "PDF");
        if (MarkdownExtensions.Contains(extension)) return (ICON_MARKDOWN, "Markdown");
        return (0, "?");
    }

    private void BtnRemoveSelected_Click(object? sender, EventArgs e)
    {
        var indices = listViewFiles.SelectedIndices.Cast<int>().OrderByDescending(i => i).ToList();
        foreach (var index in indices)
        {
            filePaths.RemoveAt(index);
            listViewFiles.Items.RemoveAt(index);
        }
        UpdateTitle();
        UpdatePreview();
    }

    private void BtnMoveUp_Click(object? sender, EventArgs e)
    {
        if (listViewFiles.SelectedIndices.Count > 0 && listViewFiles.SelectedIndices[0] > 0)
        {
            int index = listViewFiles.SelectedIndices[0];
            SwapItems(index, index - 1);
            listViewFiles.Items[index - 1].Selected = true;
            listViewFiles.Items[index - 1].Focused = true;
        }
    }

    private void BtnMoveDown_Click(object? sender, EventArgs e)
    {
        if (listViewFiles.SelectedIndices.Count > 0 && listViewFiles.SelectedIndices[0] < listViewFiles.Items.Count - 1)
        {
            int index = listViewFiles.SelectedIndices[0];
            SwapItems(index, index + 1);
            listViewFiles.Items[index + 1].Selected = true;
            listViewFiles.Items[index + 1].Focused = true;
        }
    }

    private void SwapItems(int index1, int index2)
    {
        (filePaths[index1], filePaths[index2]) = (filePaths[index2], filePaths[index1]);

        var item1 = listViewFiles.Items[index1];
        var item2 = listViewFiles.Items[index2];

        var text1 = item1.Text;
        var sub1 = item1.SubItems[1].Text;
        var icon1 = item1.ImageIndex;

        item1.Text = item2.Text;
        item1.SubItems[1].Text = item2.SubItems[1].Text;
        item1.ImageIndex = item2.ImageIndex;

        item2.Text = text1;
        item2.SubItems[1].Text = sub1;
        item2.ImageIndex = icon1;
    }

    private void BtnClear_Click(object? sender, EventArgs e)
    {
        filePaths.Clear();
        listViewFiles.Items.Clear();
        UpdateTitle();
        UpdatePreview();
    }

    private void UpdateTitle()
    {
        this.Text = $"PDF Merger Standard - {filePaths.Count} fichier(s)";
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
        btnPreviewResult.Enabled = enabled;
        btnTogglePreview.Enabled = enabled;
    }

    private async void BtnPreviewResult_Click(object? sender, EventArgs e)
    {
        if (filePaths.Count == 0)
        {
            MessageBox.Show("Veuillez ajouter au moins un fichier.", "Aucun fichier",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var tempPath = Path.Combine(Path.GetTempPath(), $"PDFMerger_Preview_{Guid.NewGuid():N}.pdf");

        progressBar.Visible = true;
        progressBar.Value = 0;
        progressBar.Maximum = filePaths.Count;
        SetButtonsEnabled(false);

        try
        {
            await Task.Run(() => CreatePdf(tempPath));

            // Ouvrir le PDF avec l'application par défaut
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = tempPath,
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(psi);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Erreur lors de la création de l'aperçu:\n{ex.Message}", "Erreur",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            progressBar.Visible = false;
            SetButtonsEnabled(true);
        }
    }

    private void CreatePdf(string outputPath)
    {
        using var document = new PdfSharpDocument();
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

    private async void PdfToWord_Click(object? sender, EventArgs e)
    {
        using var openDialog = new OpenFileDialog
        {
            Title = "Sélectionner un fichier PDF",
            Filter = "PDF|*.pdf",
            Multiselect = false
        };

        if (openDialog.ShowDialog() != DialogResult.OK)
            return;

        using var saveDialog = new SaveFileDialog
        {
            Title = "Enregistrer le fichier Word",
            Filter = "Document Word|*.docx",
            DefaultExt = "docx",
            FileName = Path.GetFileNameWithoutExtension(openDialog.FileName) + ".docx"
        };

        if (saveDialog.ShowDialog() != DialogResult.OK)
            return;

        progressBar.Visible = true;
        progressBar.Style = ProgressBarStyle.Marquee;
        SetButtonsEnabled(false);

        try
        {
            await Task.Run(() => ConvertPdfToWord(openDialog.FileName, saveDialog.FileName));

            MessageBox.Show($"Conversion réussie!\n{saveDialog.FileName}", "Succès",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Erreur lors de la conversion:\n{ex.Message}", "Erreur",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            progressBar.Style = ProgressBarStyle.Blocks;
            progressBar.Visible = false;
            SetButtonsEnabled(true);
        }
    }

    private static void ConvertPdfToWord(string pdfPath, string wordPath)
    {
        // Extraire le texte du PDF avec iText7
        var text = new System.Text.StringBuilder();

        using (var pdfReader = new iText.Kernel.Pdf.PdfReader(pdfPath))
        using (var pdfDoc = new iText.Kernel.Pdf.PdfDocument(pdfReader))
        {
            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                var page = pdfDoc.GetPage(i);
                var strategy = new SimpleTextExtractionStrategy();
                var pageText = PdfTextExtractor.GetTextFromPage(page, strategy);
                text.AppendLine(pageText);
                text.AppendLine(); // Séparateur de page
            }
        }

        // Créer le document Word avec OpenXML
        using var wordDoc = WordprocessingDocument.Create(wordPath, WordprocessingDocumentType.Document);

        var mainPart = wordDoc.AddMainDocumentPart();
        mainPart.Document = new WordDocument();
        var body = mainPart.Document.AppendChild(new WordBody());

        // Ajouter le texte paragraphe par paragraphe
        var paragraphs = text.ToString().Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

        foreach (var para in paragraphs)
        {
            var paragraph = body.AppendChild(new WordParagraph());
            var run = paragraph.AppendChild(new WordRun());
            run.AppendChild(new WordText(para) { Space = SpaceProcessingModeValues.Preserve });
        }
    }

    private void OpenPdfSam_Click(object? sender, EventArgs e)
    {
        // Chemins possibles pour PDFsam
        var possiblePaths = new[]
        {
            @"C:\Program Files\PDFsam Basic\pdfsam.exe",
            @"C:\Program Files\PDFsam Basic\PDFsam.exe",
            @"C:\Program Files (x86)\PDFsam Basic\pdfsam.exe",
            @"C:\Program Files (x86)\PDFsam Basic\PDFsam.exe",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"PDFsam Basic\pdfsam.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\PDFsam Basic\pdfsam.exe"),
            @"C:\Program Files\PDFsam\pdfsam.exe",
            @"C:\Program Files (x86)\PDFsam\pdfsam.exe",
            // Version Enhanced
            @"C:\Program Files\PDFsam Enhanced\pdfsam.exe",
            @"C:\Program Files (x86)\PDFsam Enhanced\pdfsam.exe",
        };

        string? pdfSamPath = null;

        foreach (var path in possiblePaths)
        {
            if (File.Exists(path))
            {
                pdfSamPath = path;
                break;
            }
        }

        if (pdfSamPath == null)
        {
            var result = MessageBox.Show(
                "PDFsam n'a pas été trouvé sur votre système.\n\n" +
                "Voulez-vous le télécharger depuis le site officiel?",
                "PDFsam non trouvé",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "https://pdfsam.org/download-pdfsam-basic/",
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(psi);
            }
            return;
        }

        try
        {
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = pdfSamPath,
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(psi);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Erreur lors du lancement de PDF SAM:\n{ex.Message}", "Erreur",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
