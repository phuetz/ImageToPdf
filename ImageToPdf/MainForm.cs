using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;

namespace ImageToPdf;

public class MainForm : Form
{
    private ListBox listBoxImages;
    private Button btnAddImages;
    private Button btnRemoveSelected;
    private Button btnMoveUp;
    private Button btnMoveDown;
    private Button btnClear;
    private Button btnConvert;
    private Label lblInfo;
    private ProgressBar progressBar;

    private List<string> imagePaths = new();

    public MainForm()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        this.Text = "Image to PDF Converter";
        this.Size = new Size(600, 500);
        this.MinimumSize = new Size(500, 400);
        this.StartPosition = FormStartPosition.CenterScreen;

        // Label info
        lblInfo = new Label
        {
            Text = "Sélectionnez des images à convertir en PDF:",
            Location = new Point(12, 12),
            Size = new Size(350, 20)
        };

        // ListBox pour les images
        listBoxImages = new ListBox
        {
            Location = new Point(12, 35),
            Size = new Size(400, 350),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            SelectionMode = SelectionMode.MultiExtended,
            HorizontalScrollbar = true
        };

        // Boutons
        int btnX = 420;
        int btnWidth = 150;

        btnAddImages = new Button
        {
            Text = "Ajouter images...",
            Location = new Point(btnX, 35),
            Size = new Size(btnWidth, 30),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        btnAddImages.Click += BtnAddImages_Click;

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
            Text = "Convertir en PDF",
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
            Size = new Size(556, 25),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            Visible = false
        };

        // Ajouter les contrôles
        this.Controls.AddRange(new Control[]
        {
            lblInfo, listBoxImages, btnAddImages, btnRemoveSelected,
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
            AddImages(files);
        }
    }

    private void BtnAddImages_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Sélectionner des images",
            Filter = "Images|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif|Tous les fichiers|*.*",
            Multiselect = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            AddImages(dialog.FileNames);
        }
    }

    private void AddImages(string[] files)
    {
        var validExtensions = new[] { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif" };

        foreach (var file in files)
        {
            var ext = Path.GetExtension(file).ToLowerInvariant();
            if (validExtensions.Contains(ext) && !imagePaths.Contains(file))
            {
                imagePaths.Add(file);
                listBoxImages.Items.Add(Path.GetFileName(file));
            }
        }

        UpdateTitle();
    }

    private void BtnRemoveSelected_Click(object? sender, EventArgs e)
    {
        var indices = listBoxImages.SelectedIndices.Cast<int>().OrderByDescending(i => i).ToList();
        foreach (var index in indices)
        {
            imagePaths.RemoveAt(index);
            listBoxImages.Items.RemoveAt(index);
        }
        UpdateTitle();
    }

    private void BtnMoveUp_Click(object? sender, EventArgs e)
    {
        if (listBoxImages.SelectedIndex > 0)
        {
            int index = listBoxImages.SelectedIndex;
            SwapItems(index, index - 1);
            listBoxImages.SelectedIndex = index - 1;
        }
    }

    private void BtnMoveDown_Click(object? sender, EventArgs e)
    {
        if (listBoxImages.SelectedIndex >= 0 && listBoxImages.SelectedIndex < listBoxImages.Items.Count - 1)
        {
            int index = listBoxImages.SelectedIndex;
            SwapItems(index, index + 1);
            listBoxImages.SelectedIndex = index + 1;
        }
    }

    private void SwapItems(int index1, int index2)
    {
        (imagePaths[index1], imagePaths[index2]) = (imagePaths[index2], imagePaths[index1]);

        var temp = listBoxImages.Items[index1];
        listBoxImages.Items[index1] = listBoxImages.Items[index2];
        listBoxImages.Items[index2] = temp;
    }

    private void BtnClear_Click(object? sender, EventArgs e)
    {
        imagePaths.Clear();
        listBoxImages.Items.Clear();
        UpdateTitle();
    }

    private void UpdateTitle()
    {
        this.Text = $"Image to PDF Converter - {imagePaths.Count} image(s)";
    }

    private async void BtnConvert_Click(object? sender, EventArgs e)
    {
        if (imagePaths.Count == 0)
        {
            MessageBox.Show("Veuillez ajouter au moins une image.", "Aucune image",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var saveDialog = new SaveFileDialog
        {
            Title = "Enregistrer le PDF",
            Filter = "PDF|*.pdf",
            DefaultExt = "pdf",
            FileName = "images.pdf"
        };

        if (saveDialog.ShowDialog() != DialogResult.OK)
            return;

        progressBar.Visible = true;
        progressBar.Value = 0;
        progressBar.Maximum = imagePaths.Count;
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
        btnAddImages.Enabled = enabled;
        btnRemoveSelected.Enabled = enabled;
        btnMoveUp.Enabled = enabled;
        btnMoveDown.Enabled = enabled;
        btnClear.Enabled = enabled;
        btnConvert.Enabled = enabled;
    }

    private void CreatePdf(string outputPath)
    {
        using var document = new PdfDocument();
        document.Info.Title = "Images converties en PDF";
        document.Info.Creator = "Image to PDF Converter";

        for (int i = 0; i < imagePaths.Count; i++)
        {
            var imagePath = imagePaths[i];

            try
            {
                using var stream = File.OpenRead(imagePath);
                var image = XImage.FromStream(() => stream);

                // Créer une page avec la taille de l'image
                var page = document.AddPage();
                page.Width = XUnit.FromPoint(image.PointWidth);
                page.Height = XUnit.FromPoint(image.PointHeight);

                using var gfx = XGraphics.FromPdfPage(page);
                gfx.DrawImage(image, 0, 0, page.Width, page.Height);
            }
            catch (Exception ex)
            {
                // Log l'erreur mais continue avec les autres images
                System.Diagnostics.Debug.WriteLine($"Erreur avec {imagePath}: {ex.Message}");
            }

            // Mettre à jour la barre de progression
            this.Invoke(() =>
            {
                progressBar.Value = i + 1;
            });
        }

        document.Save(outputPath);
    }
}
