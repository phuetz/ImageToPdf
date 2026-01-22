using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf.IO;
using System.Drawing.Drawing2D;
using PdfSharpDocument = PdfSharpCore.Pdf.PdfDocument;

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
    private Label lblInfo = null!;
    private ProgressBar progressBar = null!;

    private Panel mainPanel = null!;
    private Panel buttonPanel = null!;
    private Panel contentPanel = null!;

    private List<string> filePaths = new();

    private static readonly string[] ImageExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif" };
    private static readonly string[] PdfExtensions = { ".pdf" };

    private const int ICON_IMAGE = 0;
    private const int ICON_PDF = 1;

    public MainForm()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        this.Text = "PDF Merger Lite";
        this.Size = new Size(550, 450);
        this.MinimumSize = new Size(450, 380);
        this.StartPosition = FormStartPosition.CenterScreen;

        CreateImageList();
        CreateMainPanel();

        this.Controls.Add(mainPanel);

        this.AllowDrop = true;
        this.DragEnter += MainForm_DragEnter;
        this.DragDrop += MainForm_DragDrop;
    }

    private void CreateMainPanel()
    {
        mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(10)
        };

        lblInfo = new Label
        {
            Text = "Ajoutez des images ou PDF à fusionner:",
            Dock = DockStyle.Top,
            Height = 25,
            Padding = new Padding(0, 5, 0, 0)
        };

        buttonPanel = new Panel
        {
            Dock = DockStyle.Right,
            Width = 130,
            Padding = new Padding(10, 0, 0, 0)
        };

        int btnWidth = 115;
        int btnY = 0;
        int btnSpacing = 35;

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
        btnY += btnSpacing + 15;

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
        btnY += btnSpacing + 15;

        btnClear = new Button
        {
            Text = "Tout effacer",
            Location = new Point(10, btnY),
            Size = new Size(btnWidth, 28)
        };
        btnClear.Click += BtnClear_Click;
        btnY += btnSpacing + 25;

        btnConvert = new Button
        {
            Text = "Créer le PDF",
            Location = new Point(10, btnY),
            Size = new Size(btnWidth, 40),
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font(this.Font.FontFamily, 9, FontStyle.Bold)
        };
        btnConvert.Click += BtnConvert_Click;

        buttonPanel.Controls.AddRange(new Control[]
        {
            btnAddFiles, btnRemoveSelected, btnMoveUp, btnMoveDown,
            btnClear, btnConvert
        });

        contentPanel = new Panel
        {
            Dock = DockStyle.Fill
        };

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
        listViewFiles.Columns.Add("Type", 60);

        progressBar = new ProgressBar
        {
            Dock = DockStyle.Bottom,
            Height = 23,
            Visible = false
        };

        contentPanel.Controls.Add(listViewFiles);
        contentPanel.Controls.Add(progressBar);

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
    }

    private static Bitmap CreateImageIcon()
    {
        var bmp = new Bitmap(20, 20);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);

        using var framePen = new Pen(Color.FromArgb(100, 100, 100), 1);
        using var frameBrush = new SolidBrush(Color.FromArgb(240, 240, 240));
        g.FillRectangle(frameBrush, 1, 1, 18, 18);
        g.DrawRectangle(framePen, 1, 1, 17, 17);

        using var skyBrush = new SolidBrush(Color.FromArgb(135, 206, 250));
        g.FillRectangle(skyBrush, 2, 2, 16, 10);

        using var sunBrush = new SolidBrush(Color.FromArgb(255, 200, 50));
        g.FillEllipse(sunBrush, 12, 3, 5, 5);

        var mountain = new Point[] { new(2, 17), new(8, 8), new(14, 17) };
        using var mountainBrush = new SolidBrush(Color.FromArgb(76, 175, 80));
        g.FillPolygon(mountainBrush, mountain);

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

        var docPoints = new Point[] { new(3, 1), new(14, 1), new(17, 4), new(17, 18), new(3, 18) };
        using var docBrush = new SolidBrush(Color.White);
        using var docPen = new Pen(Color.FromArgb(200, 50, 50), 1);
        g.FillPolygon(docBrush, docPoints);
        g.DrawPolygon(docPen, docPoints);

        var foldPoints = new Point[] { new(14, 1), new(14, 4), new(17, 4) };
        using var foldBrush = new SolidBrush(Color.FromArgb(230, 230, 230));
        g.FillPolygon(foldBrush, foldPoints);
        g.DrawPolygon(docPen, foldPoints);

        using var pdfBrush = new SolidBrush(Color.FromArgb(220, 50, 50));
        g.FillRectangle(pdfBrush, 3, 10, 14, 7);

        using var font = new Font("Arial", 6, FontStyle.Bold);
        using var textBrush = new SolidBrush(Color.White);
        g.DrawString("PDF", font, textBrush, 4, 10);

        return bmp;
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
            Filter = "Fichiers supportés|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif;*.pdf|" +
                     "Images|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif|" +
                     "PDF|*.pdf|" +
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
        var allValidExtensions = ImageExtensions.Concat(PdfExtensions).ToArray();

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
    }

    private void UpdateTitle()
    {
        this.Text = $"PDF Merger Lite - {filePaths.Count} fichier(s)";
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
        using var document = new PdfSharpDocument();
        document.Info.Title = "Document fusionné";
        document.Info.Creator = "PDF Merger Lite";

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
}
