using PdfSharpCore.Pdf.IO;
using System.Drawing.Drawing2D;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
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
    private ToolStrip toolStrip = null!;
    private ToolStripButton btnAddFiles = null!;
    private ToolStripButton btnRemoveSelected = null!;
    private ToolStripButton btnMoveUp = null!;
    private ToolStripButton btnMoveDown = null!;
    private ToolStripButton btnClear = null!;
    private ToolStripButton btnConvert = null!;
    private ToolStripButton btnPreviewResult = null!;
    private ToolStripButton btnTogglePreview = null!;
    private ToolStripProgressBar progressBar = null!;
    private StatusStrip statusStrip = null!;
    private ToolStripStatusLabel statusLabel = null!;

    // Panels pour la disposition
    private Panel mainPanel = null!;
    private Panel contentPanel = null!;

    // Panneau de prévisualisation
    private Panel previewPanel = null!;
    private PictureBox pictureBoxPreview = null!;
    private RichTextBox textBoxPreview = null!;
    private Label lblPreviewInfo = null!;
    private Label lblNoPreview = null!;
    private Splitter splitter = null!;

    // Contrôles de navigation et zoom pour l'aperçu
    private Panel previewControlsPanel = null!;
    private Button btnPrevPage = null!;
    private Button btnNextPage = null!;
    private Label lblPageInfo = null!;
    private TrackBar zoomTrackBar = null!;
    private Label lblZoom = null!;
    private int currentPreviewPage = 0;
    private int totalPreviewPages = 0;
    private float currentZoom = 1.0f;
    private string? currentPreviewFile = null;

    // Miniatures
    private ImageList imageListThumbnails = null!;

    // Historique récent
    private List<string> recentFiles = new();
    private const int MaxRecentFiles = 10;
    private static readonly string RecentFilesPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "PDFMerger", "recent.txt");

    // Configuration
    private static readonly string ConfigFilePath = Path.Combine(
        AppContext.BaseDirectory, "pdfmerger.conf");
    private bool toolsEnabled = false;

    // Tri
    private int sortColumn = -1;
    private SortOrder sortOrder = SortOrder.None;

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
        LoadConfiguration();
        InitializeComponent();
        LoadRecentFiles();
    }

    private void LoadConfiguration()
    {
        toolsEnabled = false;

        if (File.Exists(ConfigFilePath))
        {
            try
            {
                var lines = File.ReadAllLines(ConfigFilePath);
                foreach (var line in lines)
                {
                    var trimmed = line.Trim();
                    if (trimmed.StartsWith("#") || string.IsNullOrEmpty(trimmed))
                        continue;

                    var parts = trimmed.Split('=', 2);
                    if (parts.Length == 2)
                    {
                        var key = parts[0].Trim().ToLowerInvariant();
                        var value = parts[1].Trim().ToLowerInvariant();

                        if (key == "tools_enabled" || key == "enable_tools")
                        {
                            toolsEnabled = value == "true" || value == "1" || value == "yes";
                        }
                    }
                }
            }
            catch
            {
                toolsEnabled = false;
            }
        }
    }

    private void InitializeComponent()
    {
        this.Text = "PDF Merger";
        this.Size = new Size(1000, 650);
        this.MinimumSize = new Size(800, 550);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.White;
        this.KeyPreview = true;
        this.KeyDown += MainForm_KeyDown;

        CreateImageList();
        CreateThumbnailImageList();
        CreateCommandBar();
        CreateStatusStrip();
        CreatePreviewPanel();
        CreateMainPanel();

        // Splitter entre main et preview
        splitter = new Splitter
        {
            Dock = DockStyle.Right,
            Width = 6,
            BackColor = Color.FromArgb(218, 218, 218),
            MinExtra = 300,
            MinSize = 250,
            Visible = false,
            Cursor = Cursors.VSplit
        };
        splitter.Paint += (s, e) =>
        {
            // Dessiner une poignée au centre du splitter
            var rect = splitter.ClientRectangle;
            int centerY = rect.Height / 2;
            using var brush = new SolidBrush(Color.FromArgb(160, 160, 160));
            for (int i = -12; i <= 12; i += 6)
            {
                e.Graphics.FillEllipse(brush, 2, centerY + i, 2, 2);
            }
        };

        // Ordre d'ajout important pour le docking (inverse de l'ordre visuel)
        this.Controls.Add(previewPanel);  // À droite
        this.Controls.Add(splitter);       // Splitter
        this.Controls.Add(mainPanel);      // Fill - reste de l'espace
        this.Controls.Add(statusStrip);
        this.Controls.Add(toolStrip);

        // Support du drag & drop
        this.AllowDrop = true;
        this.DragEnter += MainForm_DragEnter;
        this.DragDrop += MainForm_DragDrop;
    }

    private void CreateCommandBar()
    {
        toolStrip = new ToolStrip
        {
            GripStyle = ToolStripGripStyle.Hidden,
            BackColor = Color.FromArgb(249, 249, 249),
            Padding = new Padding(8, 6, 8, 6),
            ImageScalingSize = new Size(24, 24),
            RenderMode = ToolStripRenderMode.System,
            AutoSize = false,
            Height = 52
        };

        // Style Windows 11 - boutons avec icônes 24x24
        btnAddFiles = CreateCommandButton("Nouveau", CreateAddIcon24(), "Ajouter des fichiers");
        btnAddFiles.Click += BtnAddFiles_Click;

        btnRemoveSelected = CreateCommandButton("Supprimer", CreateDeleteIcon24(), "Supprimer la sélection");
        btnRemoveSelected.Click += BtnRemoveSelected_Click;

        btnMoveUp = CreateCommandButton("", CreateUpIcon24(), "Monter");
        btnMoveUp.DisplayStyle = ToolStripItemDisplayStyle.Image;
        btnMoveUp.Click += BtnMoveUp_Click;

        btnMoveDown = CreateCommandButton("", CreateDownIcon24(), "Descendre");
        btnMoveDown.DisplayStyle = ToolStripItemDisplayStyle.Image;
        btnMoveDown.Click += BtnMoveDown_Click;

        btnClear = CreateCommandButton("", CreateClearIcon24(), "Tout effacer");
        btnClear.DisplayStyle = ToolStripItemDisplayStyle.Image;
        btnClear.Click += BtnClear_Click;

        btnTogglePreview = CreateCommandButton("", CreatePreviewIcon24(), "Volet de visualisation");
        btnTogglePreview.DisplayStyle = ToolStripItemDisplayStyle.Image;
        btnTogglePreview.CheckOnClick = true;
        btnTogglePreview.Click += BtnTogglePreview_Click;

        btnPreviewResult = CreateCommandButton("Aperçu", CreateViewIcon24(), "Voir le résultat");
        btnPreviewResult.Click += BtnPreviewResult_Click;

        btnConvert = CreateCommandButton("Créer PDF", CreateConvertIcon24(), "Créer le fichier PDF");
        btnConvert.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        btnConvert.ForeColor = Color.FromArgb(0, 95, 184);
        btnConvert.Click += BtnConvert_Click;

        // Menu déroulant Outils
        var toolsDropDown = new ToolStripDropDownButton
        {
            Text = "Outils",
            DisplayStyle = ToolStripItemDisplayStyle.Text,
            Font = new Font("Segoe UI", 9),
            Padding = new Padding(4, 0, 4, 0)
        };
        toolsDropDown.DropDownItems.Add("PDF → Word", null, PdfToWord_Click);
        toolsDropDown.DropDownItems.Add(new ToolStripSeparator());
        toolsDropDown.DropDownItems.Add("Ouvrir PDFsam", null, OpenPdfSam_Click);

        // Menu déroulant Récents
        var recentDropDown = new ToolStripDropDownButton
        {
            Text = "Récents",
            DisplayStyle = ToolStripItemDisplayStyle.Text,
            Font = new Font("Segoe UI", 9),
            Padding = new Padding(4, 0, 4, 0)
        };
        recentDropDown.DropDownOpening += (s, e) => UpdateRecentMenu(recentDropDown);

        // Menu déroulant Vue
        var viewDropDown = new ToolStripDropDownButton
        {
            Text = "Affichage",
            Image = CreateViewModeIcon24(),
            DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
            Font = new Font("Segoe UI", 9),
            Padding = new Padding(4, 0, 4, 0)
        };

        var viewExtraLargeIcons = new ToolStripMenuItem("Très grandes icônes", null, (s, e) => SetViewMode(View.LargeIcon, 96));
        var viewLargeIcons = new ToolStripMenuItem("Grandes icônes", null, (s, e) => SetViewMode(View.LargeIcon, 64));
        var viewMediumIcons = new ToolStripMenuItem("Icônes moyennes", null, (s, e) => SetViewMode(View.LargeIcon, 48));
        var viewSmallIcons = new ToolStripMenuItem("Petites icônes", null, (s, e) => SetViewMode(View.SmallIcon, 20));
        var viewList = new ToolStripMenuItem("Liste", null, (s, e) => SetViewMode(View.List, 20));
        var viewDetails = new ToolStripMenuItem("Détails", null, (s, e) => SetViewMode(View.Details, 20));
        var viewTile = new ToolStripMenuItem("Mosaïque", null, (s, e) => SetViewMode(View.Tile, 48));

        viewDropDown.DropDownItems.AddRange(new ToolStripItem[]
        {
            viewExtraLargeIcons, viewLargeIcons, viewMediumIcons, viewSmallIcons,
            new ToolStripSeparator(),
            viewList, viewDetails, viewTile
        });

        toolStrip.Items.Add(btnAddFiles);
        toolStrip.Items.Add(CreateSeparator());
        toolStrip.Items.Add(btnRemoveSelected);
        toolStrip.Items.Add(CreateSeparator());
        toolStrip.Items.Add(btnMoveUp);
        toolStrip.Items.Add(btnMoveDown);
        toolStrip.Items.Add(CreateSeparator());
        toolStrip.Items.Add(btnClear);
        toolStrip.Items.Add(new ToolStripSeparator());
        toolStrip.Items.Add(btnTogglePreview);
        toolStrip.Items.Add(viewDropDown);
        toolStrip.Items.Add(new ToolStripSeparator());

        // Menu Outils uniquement si activé dans la configuration
        if (toolsEnabled)
        {
            toolStrip.Items.Add(toolsDropDown);
        }

        toolStrip.Items.Add(recentDropDown);
        toolStrip.Items.Add(new ToolStripSeparator());
        toolStrip.Items.Add(btnPreviewResult);
        toolStrip.Items.Add(btnConvert);
    }

    private void SetViewMode(View view, int iconSize)
    {
        listViewFiles.BeginUpdate();

        // Recréer les ImageLists avec la nouvelle taille si nécessaire
        if (iconSize != imageListThumbnails.ImageSize.Width && iconSize > 20)
        {
            RegenerateThumbnails(iconSize);
        }

        listViewFiles.View = view;

        if (view == View.Tile)
        {
            listViewFiles.TileSize = new Size(280, iconSize + 10);
        }

        listViewFiles.EndUpdate();
    }

    private void RegenerateThumbnails(int size)
    {
        imageListThumbnails.Images.Clear();
        imageListThumbnails.ImageSize = new Size(size, size);

        for (int i = 0; i < filePaths.Count; i++)
        {
            var thumbKey = $"thumb_{i + 1}";
            var thumb = CreateThumbnailWithSize(filePaths[i], size);
            imageListThumbnails.Images.Add(thumbKey, thumb);
            listViewFiles.Items[i].ImageKey = thumbKey;
        }
    }

    private Bitmap CreateThumbnailWithSize(string filePath, int size)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var thumb = new Bitmap(size, size);

        using (var g = Graphics.FromImage(thumb))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.Clear(Color.White);

            try
            {
                if (ImageExtensions.Contains(ext))
                {
                    using var stream = File.OpenRead(filePath);
                    using var img = Image.FromStream(stream);
                    var scale = Math.Min((size - 2f) / img.Width, (size - 2f) / img.Height);
                    var w = (int)(img.Width * scale);
                    var h = (int)(img.Height * scale);
                    g.DrawImage(img, (size - w) / 2, (size - h) / 2, w, h);
                }
                else if (PdfExtensions.Contains(ext))
                {
                    DrawPdfIconScaled(g, size);
                }
                else if (MarkdownExtensions.Contains(ext))
                {
                    DrawMarkdownIconScaled(g, size);
                }
            }
            catch
            {
                using var pen = new Pen(Color.Gray, 1);
                g.DrawRectangle(pen, 1, 1, size - 3, size - 3);
            }

            using var borderPen = new Pen(Color.FromArgb(200, 200, 200), 1);
            g.DrawRectangle(borderPen, 0, 0, size - 1, size - 1);
        }

        return thumb;
    }

    private static void DrawPdfIconScaled(Graphics g, int size)
    {
        var scale = size / 48f;
        using var docBrush = new SolidBrush(Color.White);
        using var docPen = new Pen(Color.FromArgb(200, 50, 50), 2 * scale);
        g.FillRectangle(docBrush, 6 * scale, 2 * scale, 36 * scale, 44 * scale);
        g.DrawRectangle(docPen, 6 * scale, 2 * scale, 36 * scale, 44 * scale);

        using var pdfBrush = new SolidBrush(Color.FromArgb(220, 50, 50));
        g.FillRectangle(pdfBrush, 8 * scale, 24 * scale, 32 * scale, 18 * scale);

        using var font = new Font("Arial", 12 * scale, FontStyle.Bold);
        using var textBrush = new SolidBrush(Color.White);
        g.DrawString("PDF", font, textBrush, 10 * scale, 27 * scale);
    }

    private static void DrawMarkdownIconScaled(Graphics g, int size)
    {
        var scale = size / 48f;
        using var docBrush = new SolidBrush(Color.White);
        using var docPen = new Pen(Color.FromArgb(50, 50, 50), 2 * scale);
        g.FillRectangle(docBrush, 6 * scale, 2 * scale, 36 * scale, 44 * scale);
        g.DrawRectangle(docPen, 6 * scale, 2 * scale, 36 * scale, 44 * scale);

        using var mdBrush = new SolidBrush(Color.FromArgb(33, 150, 243));
        g.FillRectangle(mdBrush, 10 * scale, 14 * scale, 28 * scale, 20 * scale);

        using var font = new Font("Arial", 10 * scale, FontStyle.Bold);
        using var textBrush = new SolidBrush(Color.White);
        g.DrawString("MD", font, textBrush, 14 * scale, 18 * scale);
    }

    private static Bitmap CreateViewModeIcon24()
    {
        var bmp = new Bitmap(24, 24);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        using var brush = new SolidBrush(Color.FromArgb(96, 96, 96));
        // Grille 2x2
        g.FillRectangle(brush, 3, 3, 8, 8);
        g.FillRectangle(brush, 13, 3, 8, 8);
        g.FillRectangle(brush, 3, 13, 8, 8);
        g.FillRectangle(brush, 13, 13, 8, 8);
        return bmp;
    }

    private void UpdateRecentMenu(ToolStripDropDownButton menu)
    {
        menu.DropDownItems.Clear();

        if (recentFiles.Count == 0)
        {
            var emptyItem = new ToolStripMenuItem("(Aucun fichier récent)")
            {
                Enabled = false
            };
            menu.DropDownItems.Add(emptyItem);
        }
        else
        {
            foreach (var file in recentFiles.Take(MaxRecentFiles))
            {
                var fileName = Path.GetFileName(file);
                var item = new ToolStripMenuItem(fileName)
                {
                    ToolTipText = file
                };
                item.Click += (s, e) => OpenRecentFile(file);
                menu.DropDownItems.Add(item);
            }

            menu.DropDownItems.Add(new ToolStripSeparator());
            var clearItem = new ToolStripMenuItem("Effacer l'historique");
            clearItem.Click += (s, e) =>
            {
                recentFiles.Clear();
                SaveRecentFiles();
            };
            menu.DropDownItems.Add(clearItem);
        }
    }

    private void OpenRecentFile(string filePath)
    {
        if (File.Exists(filePath))
        {
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(psi);
        }
        else
        {
            MessageBox.Show($"Le fichier n'existe plus:\n{filePath}", "Fichier introuvable",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            recentFiles.Remove(filePath);
            SaveRecentFiles();
        }
    }

    private static ToolStripButton CreateCommandButton(string text, Image icon, string tooltip)
    {
        return new ToolStripButton
        {
            Text = text,
            Image = icon,
            ImageAlign = ContentAlignment.MiddleLeft,
            TextAlign = ContentAlignment.MiddleLeft,
            DisplayStyle = string.IsNullOrEmpty(text) ? ToolStripItemDisplayStyle.Image : ToolStripItemDisplayStyle.ImageAndText,
            Font = new Font("Segoe UI", 9),
            Padding = new Padding(4, 0, 4, 0),
            ToolTipText = tooltip,
            AutoSize = true
        };
    }

    private static ToolStripLabel CreateSeparator()
    {
        return new ToolStripLabel(" ") { Font = new Font("Segoe UI", 6) };
    }

    // Icônes 24x24 style Windows 11
    private static Bitmap CreateAddIcon24()
    {
        var bmp = new Bitmap(24, 24);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        using var pen = new Pen(Color.FromArgb(0, 120, 212), 2.5f);
        g.DrawLine(pen, 12, 4, 12, 20);
        g.DrawLine(pen, 4, 12, 20, 12);
        return bmp;
    }

    private static Bitmap CreateDeleteIcon24()
    {
        var bmp = new Bitmap(24, 24);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        using var pen = new Pen(Color.FromArgb(196, 43, 28), 1.5f);
        // Corbeille
        g.DrawRectangle(pen, 6, 7, 12, 14);
        g.DrawLine(pen, 4, 7, 20, 7);
        g.DrawLine(pen, 9, 4, 15, 4);
        g.DrawLine(pen, 9, 4, 9, 7);
        g.DrawLine(pen, 15, 4, 15, 7);
        g.DrawLine(pen, 9, 10, 9, 18);
        g.DrawLine(pen, 12, 10, 12, 18);
        g.DrawLine(pen, 15, 10, 15, 18);
        return bmp;
    }

    private static Bitmap CreateUpIcon24()
    {
        var bmp = new Bitmap(24, 24);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        using var pen = new Pen(Color.FromArgb(96, 96, 96), 2.5f);
        g.DrawLine(pen, 12, 5, 12, 19);
        g.DrawLine(pen, 6, 11, 12, 5);
        g.DrawLine(pen, 18, 11, 12, 5);
        return bmp;
    }

    private static Bitmap CreateDownIcon24()
    {
        var bmp = new Bitmap(24, 24);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        using var pen = new Pen(Color.FromArgb(96, 96, 96), 2.5f);
        g.DrawLine(pen, 12, 5, 12, 19);
        g.DrawLine(pen, 6, 13, 12, 19);
        g.DrawLine(pen, 18, 13, 12, 19);
        return bmp;
    }

    private static Bitmap CreateClearIcon24()
    {
        var bmp = new Bitmap(24, 24);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        using var pen = new Pen(Color.FromArgb(96, 96, 96), 2f);
        g.DrawLine(pen, 6, 6, 18, 18);
        g.DrawLine(pen, 18, 6, 6, 18);
        return bmp;
    }

    private static Bitmap CreatePreviewIcon24()
    {
        var bmp = new Bitmap(24, 24);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        using var pen = new Pen(Color.FromArgb(96, 96, 96), 1.5f);
        // Deux panneaux
        g.DrawRectangle(pen, 3, 3, 9, 18);
        g.DrawRectangle(pen, 13, 3, 8, 18);
        using var brush = new SolidBrush(Color.FromArgb(0, 120, 212));
        g.FillRectangle(brush, 14, 4, 7, 17);
        return bmp;
    }

    private static Bitmap CreateViewIcon24()
    {
        var bmp = new Bitmap(24, 24);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        using var brush = new SolidBrush(Color.FromArgb(0, 120, 212));
        // Oeil
        g.FillEllipse(brush, 2, 7, 20, 10);
        using var whiteBrush = new SolidBrush(Color.White);
        g.FillEllipse(whiteBrush, 8, 9, 8, 6);
        using var blackBrush = new SolidBrush(Color.FromArgb(0, 90, 158));
        g.FillEllipse(blackBrush, 10, 10, 4, 4);
        return bmp;
    }

    private static Bitmap CreateConvertIcon24()
    {
        var bmp = new Bitmap(24, 24);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        using var brush = new SolidBrush(Color.FromArgb(196, 43, 28));
        g.FillRectangle(brush, 3, 3, 18, 18);
        using var font = new Font("Arial", 8, FontStyle.Bold);
        using var textBrush = new SolidBrush(Color.White);
        g.DrawString("PDF", font, textBrush, 3, 7);
        return bmp;
    }

    private void CreateStatusStrip()
    {
        statusStrip = new StatusStrip
        {
            BackColor = Color.FromArgb(249, 249, 249),
            SizingGrip = false
        };

        statusLabel = new ToolStripStatusLabel
        {
            Text = "Glissez des fichiers ici ou cliquez sur Nouveau",
            Spring = true,
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(96, 96, 96)
        };

        progressBar = new ToolStripProgressBar
        {
            Width = 150,
            Visible = false
        };

        statusStrip.Items.Add(statusLabel);
        statusStrip.Items.Add(progressBar);
    }

    private void CreateMainPanel()
    {
        mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(0)
        };

        // Panneau du contenu (liste)
        contentPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White
        };

        // ListView style Windows 11 avec miniatures
        listViewFiles = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Tile,
            FullRowSelect = true,
            SmallImageList = imageListIcons,
            LargeImageList = imageListThumbnails,
            MultiSelect = false,
            HeaderStyle = ColumnHeaderStyle.Clickable,
            BorderStyle = BorderStyle.None,
            GridLines = false,
            Font = new Font("Segoe UI", 9),
            BackColor = Color.White,
            TileSize = new Size(280, 58)
        };

        listViewFiles.Columns.Add("Nom", 300);
        listViewFiles.Columns.Add("Type", 80);
        listViewFiles.Columns.Add("Date de modification", 140);
        listViewFiles.SelectedIndexChanged += ListViewFiles_SelectedIndexChanged;
        listViewFiles.ColumnClick += ListViewFiles_ColumnClick;

        contentPanel.Controls.Add(listViewFiles);
        mainPanel.Controls.Add(contentPanel);
    }

    private void ListViewFiles_ColumnClick(object? sender, ColumnClickEventArgs e)
    {
        if (e.Column == sortColumn)
        {
            sortOrder = sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
        }
        else
        {
            sortColumn = e.Column;
            sortOrder = SortOrder.Ascending;
        }

        SortListView();
    }

    private void SortListView()
    {
        if (sortColumn < 0 || listViewFiles.Items.Count == 0) return;

        var items = new List<(ListViewItem item, string path)>();
        for (int i = 0; i < listViewFiles.Items.Count; i++)
        {
            items.Add((listViewFiles.Items[i], filePaths[i]));
        }

        items = sortColumn switch
        {
            0 => sortOrder == SortOrder.Ascending
                ? items.OrderBy(x => x.item.Text).ToList()
                : items.OrderByDescending(x => x.item.Text).ToList(),
            1 => sortOrder == SortOrder.Ascending
                ? items.OrderBy(x => x.item.SubItems[1].Text).ToList()
                : items.OrderByDescending(x => x.item.SubItems[1].Text).ToList(),
            2 => sortOrder == SortOrder.Ascending
                ? items.OrderBy(x => DateTime.TryParse(x.item.SubItems[2].Text, out var d) ? d : DateTime.MinValue).ToList()
                : items.OrderByDescending(x => DateTime.TryParse(x.item.SubItems[2].Text, out var d) ? d : DateTime.MinValue).ToList(),
            _ => items
        };

        listViewFiles.BeginUpdate();
        listViewFiles.Items.Clear();
        filePaths.Clear();

        foreach (var (item, path) in items)
        {
            listViewFiles.Items.Add(item);
            filePaths.Add(path);
        }
        listViewFiles.EndUpdate();
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

    private void CreateThumbnailImageList()
    {
        imageListThumbnails = new ImageList
        {
            ImageSize = new Size(48, 48),
            ColorDepth = ColorDepth.Depth32Bit
        };
    }

    private Bitmap CreateThumbnail(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var thumb = new Bitmap(48, 48);

        using (var g = Graphics.FromImage(thumb))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.Clear(Color.White);

            try
            {
                if (ImageExtensions.Contains(ext))
                {
                    using var stream = File.OpenRead(filePath);
                    using var img = Image.FromStream(stream);
                    var scale = Math.Min(46f / img.Width, 46f / img.Height);
                    var w = (int)(img.Width * scale);
                    var h = (int)(img.Height * scale);
                    g.DrawImage(img, (48 - w) / 2, (48 - h) / 2, w, h);
                }
                else if (PdfExtensions.Contains(ext))
                {
                    // Icône PDF agrandie
                    using var pdfIcon = CreatePdfIconLarge();
                    g.DrawImage(pdfIcon, 0, 0, 48, 48);
                }
                else if (MarkdownExtensions.Contains(ext))
                {
                    // Icône Markdown agrandie
                    using var mdIcon = CreateMarkdownIconLarge();
                    g.DrawImage(mdIcon, 0, 0, 48, 48);
                }
            }
            catch
            {
                // En cas d'erreur, dessiner un placeholder
                using var pen = new Pen(Color.Gray, 1);
                g.DrawRectangle(pen, 1, 1, 45, 45);
                g.DrawLine(pen, 1, 1, 46, 46);
                g.DrawLine(pen, 46, 1, 1, 46);
            }

            // Bordure
            using var borderPen = new Pen(Color.FromArgb(200, 200, 200), 1);
            g.DrawRectangle(borderPen, 0, 0, 47, 47);
        }

        return thumb;
    }

    private static Bitmap CreatePdfIconLarge()
    {
        var bmp = new Bitmap(48, 48);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);

        using var docBrush = new SolidBrush(Color.White);
        using var docPen = new Pen(Color.FromArgb(200, 50, 50), 2);
        g.FillRectangle(docBrush, 6, 2, 36, 44);
        g.DrawRectangle(docPen, 6, 2, 36, 44);

        using var pdfBrush = new SolidBrush(Color.FromArgb(220, 50, 50));
        g.FillRectangle(pdfBrush, 8, 24, 32, 18);

        using var font = new Font("Arial", 12, FontStyle.Bold);
        using var textBrush = new SolidBrush(Color.White);
        g.DrawString("PDF", font, textBrush, 10, 27);

        return bmp;
    }

    private static Bitmap CreateMarkdownIconLarge()
    {
        var bmp = new Bitmap(48, 48);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);

        using var docBrush = new SolidBrush(Color.White);
        using var docPen = new Pen(Color.FromArgb(50, 50, 50), 2);
        g.FillRectangle(docBrush, 6, 2, 36, 44);
        g.DrawRectangle(docPen, 6, 2, 36, 44);

        using var mdBrush = new SolidBrush(Color.FromArgb(33, 150, 243));
        g.FillRectangle(mdBrush, 10, 14, 28, 20);

        using var font = new Font("Arial", 10, FontStyle.Bold);
        using var textBrush = new SolidBrush(Color.White);
        g.DrawString("MD", font, textBrush, 14, 18);

        return bmp;
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
            BackColor = Color.FromArgb(243, 243, 243),
            Padding = new Padding(12, 8, 12, 8)
        };

        lblPreviewInfo = new Label
        {
            Text = "",
            Dock = DockStyle.Top,
            Height = 20,
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(96, 96, 96)
        };

        // Panneau de contrôles (navigation + zoom)
        previewControlsPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 70,
            BackColor = Color.FromArgb(243, 243, 243),
            Padding = new Padding(0, 8, 0, 0)
        };

        // Navigation pages
        var navPanel = new Panel { Dock = DockStyle.Top, Height = 28 };

        btnPrevPage = new Button
        {
            Text = "◀",
            Width = 32,
            Height = 26,
            Location = new Point(0, 0),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9),
            Enabled = false
        };
        btnPrevPage.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        btnPrevPage.Click += BtnPrevPage_Click;

        lblPageInfo = new Label
        {
            Text = "",
            Width = 80,
            Height = 26,
            Location = new Point(36, 0),
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 9)
        };

        btnNextPage = new Button
        {
            Text = "▶",
            Width = 32,
            Height = 26,
            Location = new Point(120, 0),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9),
            Enabled = false
        };
        btnNextPage.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        btnNextPage.Click += BtnNextPage_Click;

        navPanel.Controls.AddRange(new Control[] { btnPrevPage, lblPageInfo, btnNextPage });

        // Zoom
        var zoomPanel = new Panel { Dock = DockStyle.Bottom, Height = 30 };

        var lblZoomIcon = new Label
        {
            Text = "🔍",
            Width = 24,
            Height = 24,
            Location = new Point(0, 3),
            Font = new Font("Segoe UI", 10)
        };

        zoomTrackBar = new TrackBar
        {
            Minimum = 25,
            Maximum = 200,
            Value = 100,
            TickFrequency = 25,
            Width = 180,
            Height = 30,
            Location = new Point(24, 0),
            TickStyle = TickStyle.None
        };
        zoomTrackBar.ValueChanged += ZoomTrackBar_ValueChanged;

        lblZoom = new Label
        {
            Text = "100%",
            Width = 45,
            Height = 24,
            Location = new Point(208, 3),
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new Font("Segoe UI", 9)
        };

        zoomPanel.Controls.AddRange(new Control[] { lblZoomIcon, zoomTrackBar, lblZoom });

        previewControlsPanel.Controls.Add(zoomPanel);
        previewControlsPanel.Controls.Add(navPanel);

        var previewContentPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(1)
        };

        pictureBoxPreview = new PictureBox
        {
            Dock = DockStyle.Fill,
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.White,
            BorderStyle = BorderStyle.None,
            Visible = false
        };

        textBoxPreview = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            BackColor = Color.White,
            BorderStyle = BorderStyle.None,
            Font = new Font("Segoe UI", 9),
            Visible = false
        };

        lblNoPreview = new Label
        {
            Dock = DockStyle.Fill,
            Text = "Sélectionnez un élément\npour afficher un aperçu.",
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = Color.FromArgb(128, 128, 128),
            Font = new Font("Segoe UI", 9),
            BackColor = Color.White
        };

        previewContentPanel.Controls.Add(pictureBoxPreview);
        previewContentPanel.Controls.Add(textBoxPreview);
        previewContentPanel.Controls.Add(lblNoPreview);

        previewPanel.Controls.Add(previewContentPanel);
        previewPanel.Controls.Add(previewControlsPanel);
        previewPanel.Controls.Add(lblPreviewInfo);
    }

    private void BtnPrevPage_Click(object? sender, EventArgs e)
    {
        if (currentPreviewPage > 0)
        {
            currentPreviewPage--;
            UpdatePdfPagePreview();
        }
    }

    private void BtnNextPage_Click(object? sender, EventArgs e)
    {
        if (currentPreviewPage < totalPreviewPages - 1)
        {
            currentPreviewPage++;
            UpdatePdfPagePreview();
        }
    }

    private void ZoomTrackBar_ValueChanged(object? sender, EventArgs e)
    {
        currentZoom = zoomTrackBar.Value / 100f;
        lblZoom.Text = $"{zoomTrackBar.Value}%";

        if (pictureBoxPreview.Visible && currentPreviewFile != null)
        {
            var ext = Path.GetExtension(currentPreviewFile).ToLowerInvariant();
            if (PdfExtensions.Contains(ext))
            {
                UpdatePdfPagePreview();
            }
        }
    }

    private async void UpdatePdfPagePreview()
    {
        if (currentPreviewFile == null) return;

        try
        {
            var file = await Windows.Storage.StorageFile.GetFileFromPathAsync(currentPreviewFile);
            var pdfDoc = await Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(file);

            if (currentPreviewPage < pdfDoc.PageCount)
            {
                using var page = pdfDoc.GetPage((uint)currentPreviewPage);

                using var stream = new Windows.Storage.Streams.InMemoryRandomAccessStream();

                var baseWidth = (uint)Math.Min(pictureBoxPreview.Width * 2, 800);
                var options = new Windows.Data.Pdf.PdfPageRenderOptions
                {
                    DestinationWidth = (uint)(baseWidth * currentZoom),
                    BackgroundColor = Windows.UI.Color.FromArgb(255, 255, 255, 255)
                };

                await page.RenderToStreamAsync(stream, options);
                stream.Seek(0);

                using var netStream = stream.AsStreamForRead();
                var img = Image.FromStream(netStream);
                pictureBoxPreview.Image?.Dispose();
                pictureBoxPreview.Image = new Bitmap(img);

                pictureBoxPreview.Visible = true;
                textBoxPreview.Visible = false;
                lblNoPreview.Visible = false;
            }

            btnPrevPage.Enabled = currentPreviewPage > 0;
            btnNextPage.Enabled = currentPreviewPage < totalPreviewPages - 1;
            lblPageInfo.Text = $"{currentPreviewPage + 1} / {totalPreviewPages}";
        }
        catch { }
    }

    private void BtnTogglePreview_Click(object? sender, EventArgs e)
    {
        previewVisible = btnTogglePreview.Checked;
        previewPanel.Visible = previewVisible;
        splitter.Visible = previewVisible;

        if (previewVisible)
        {
            if (this.Width < 1000)
            {
                this.Width = 1100;
            }
            UpdatePreview();
        }
        else
        {
            if (this.Width > 900)
            {
                this.Width = 800;
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

    private async void ShowPdfPreview(string filePath)
    {
        currentPreviewFile = filePath;
        currentPreviewPage = 0;

        try
        {
            var file = await Windows.Storage.StorageFile.GetFileFromPathAsync(filePath);
            var pdfDoc = await Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(file);

            totalPreviewPages = (int)pdfDoc.PageCount;

            // Activer les contrôles de navigation
            btnPrevPage.Enabled = false;
            btnNextPage.Enabled = totalPreviewPages > 1;
            lblPageInfo.Text = $"1 / {totalPreviewPages}";

            if (pdfDoc.PageCount > 0)
            {
                using var page = pdfDoc.GetPage(0);

                using var stream = new Windows.Storage.Streams.InMemoryRandomAccessStream();

                var baseWidth = (uint)Math.Min(pictureBoxPreview.Width * 2, 800);
                var options = new Windows.Data.Pdf.PdfPageRenderOptions
                {
                    DestinationWidth = (uint)(baseWidth * currentZoom),
                    BackgroundColor = Windows.UI.Color.FromArgb(255, 255, 255, 255)
                };

                await page.RenderToStreamAsync(stream, options);
                stream.Seek(0);

                using var netStream = stream.AsStreamForRead();
                var img = Image.FromStream(netStream);
                pictureBoxPreview.Image?.Dispose();
                pictureBoxPreview.Image = new Bitmap(img);

                pictureBoxPreview.Visible = true;
                textBoxPreview.Visible = false;
                lblNoPreview.Visible = false;
            }
        }
        catch
        {
            // Fallback: afficher les métadonnées si le rendu échoue
            totalPreviewPages = 0;
            btnPrevPage.Enabled = false;
            btnNextPage.Enabled = false;
            lblPageInfo.Text = "";

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
            pictureBoxPreview.Visible = false;
            lblNoPreview.Visible = false;
        }
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

        listViewFiles.BeginUpdate();
        foreach (var file in files)
        {
            var ext = Path.GetExtension(file).ToLowerInvariant();
            if (allValidExtensions.Contains(ext) && !filePaths.Contains(file))
            {
                filePaths.Add(file);
                var (iconIndex, typeName) = GetFileTypeInfo(ext);

                var fileInfo = new FileInfo(file);
                var dateStr = fileInfo.LastWriteTime.ToString("dd/MM/yyyy HH:mm");

                // Créer la miniature (grande icône)
                var thumbKey = $"thumb_{filePaths.Count}";
                var thumb = CreateThumbnail(file);
                imageListThumbnails.Images.Add(thumbKey, thumb);

                // Créer aussi la petite icône avec la même clé
                var smallIcon = CreateSmallIcon(ext);
                imageListIcons.Images.Add(thumbKey, smallIcon);

                var item = new ListViewItem(Path.GetFileName(file))
                {
                    ImageKey = thumbKey
                };
                item.SubItems.Add(typeName);
                item.SubItems.Add(dateStr);
                listViewFiles.Items.Add(item);
            }
        }
        listViewFiles.EndUpdate();

        UpdateTitle();
    }

    private Bitmap CreateSmallIcon(string extension)
    {
        if (ImageExtensions.Contains(extension))
            return CreateImageIcon();
        if (PdfExtensions.Contains(extension))
            return CreatePdfIcon();
        if (MarkdownExtensions.Contains(extension))
            return CreateMarkdownIcon();
        return CreateImageIcon(); // Default
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
            listViewFiles.Items[index].Selected = false;
            listViewFiles.Items[index - 1].Selected = true;
            listViewFiles.Items[index - 1].Focused = true;
            listViewFiles.Focus();
            UpdatePreview();
        }
    }

    private void BtnMoveDown_Click(object? sender, EventArgs e)
    {
        if (listViewFiles.SelectedIndices.Count > 0 && listViewFiles.SelectedIndices[0] < listViewFiles.Items.Count - 1)
        {
            int index = listViewFiles.SelectedIndices[0];
            SwapItems(index, index + 1);
            listViewFiles.Items[index].Selected = false;
            listViewFiles.Items[index + 1].Selected = true;
            listViewFiles.Items[index + 1].Focused = true;
            listViewFiles.Focus();
            UpdatePreview();
        }
    }

    private void SwapItems(int index1, int index2)
    {
        (filePaths[index1], filePaths[index2]) = (filePaths[index2], filePaths[index1]);

        var item1 = listViewFiles.Items[index1];
        var item2 = listViewFiles.Items[index2];

        var text1 = item1.Text;
        var sub1 = item1.SubItems[1].Text;
        var date1 = item1.SubItems[2].Text;
        var iconKey1 = item1.ImageKey;

        item1.Text = item2.Text;
        item1.SubItems[1].Text = item2.SubItems[1].Text;
        item1.SubItems[2].Text = item2.SubItems[2].Text;
        item1.ImageKey = item2.ImageKey;

        item2.Text = text1;
        item2.SubItems[1].Text = sub1;
        item2.SubItems[2].Text = date1;
        item2.ImageKey = iconKey1;
    }

    private void BtnClear_Click(object? sender, EventArgs e)
    {
        filePaths.Clear();
        listViewFiles.Items.Clear();
        imageListThumbnails.Images.Clear();
        // Garder les icônes de base mais supprimer les icônes dynamiques
        var keysToRemove = imageListIcons.Images.Keys.Cast<string>()
            .Where(k => k.StartsWith("thumb_")).ToList();
        foreach (var key in keysToRemove)
        {
            imageListIcons.Images.RemoveByKey(key);
        }
        UpdateTitle();
        UpdatePreview();
    }

    private void UpdateTitle()
    {
        this.Text = $"PDF Merger - {filePaths.Count} fichier(s)";
        statusLabel.Text = filePaths.Count == 0
            ? "Glissez-déposez des fichiers ou cliquez sur Ajouter"
            : $"{filePaths.Count} fichier(s) dans la liste";
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

        statusLabel.Text = "Création du PDF en cours...";
        progressBar.Visible = true;
        progressBar.Value = 0;
        progressBar.Maximum = filePaths.Count;
        SetButtonsEnabled(false);

        try
        {
            await Task.Run(() => CreatePdf(saveDialog.FileName));

            AddToRecentFiles(saveDialog.FileName);
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
            UpdateTitle();
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

        statusLabel.Text = "Génération de l'aperçu...";
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
            UpdateTitle();
        }
    }

    private void CreatePdf(string outputPath)
    {
        PdfMerger.CreatePdf(filePaths, outputPath, (current, total, fileName) =>
        {
            this.Invoke(() =>
            {
                progressBar.Value = current;
            });
        });
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

        statusLabel.Text = "Conversion en cours...";
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
            statusLabel.Text = "Prêt";
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

    // Raccourcis clavier
    private void MainForm_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Control)
        {
            switch (e.KeyCode)
            {
                case Keys.O: // Ctrl+O - Ouvrir/Ajouter fichiers
                    BtnAddFiles_Click(sender, e);
                    e.Handled = true;
                    break;
                case Keys.S: // Ctrl+S - Créer PDF
                    BtnConvert_Click(sender, e);
                    e.Handled = true;
                    break;
                case Keys.P: // Ctrl+P - Aperçu
                    btnTogglePreview.Checked = !btnTogglePreview.Checked;
                    BtnTogglePreview_Click(sender, e);
                    e.Handled = true;
                    break;
            }
        }
        else
        {
            switch (e.KeyCode)
            {
                case Keys.Delete: // Supprimer sélection
                    BtnRemoveSelected_Click(sender, e);
                    e.Handled = true;
                    break;
                case Keys.F5: // Voir résultat
                    BtnPreviewResult_Click(sender, e);
                    e.Handled = true;
                    break;
            }
        }
    }

    // Historique récent
    private void LoadRecentFiles()
    {
        try
        {
            var dir = Path.GetDirectoryName(RecentFilesPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (File.Exists(RecentFilesPath))
            {
                recentFiles = File.ReadAllLines(RecentFilesPath)
                    .Where(f => File.Exists(f))
                    .Take(MaxRecentFiles)
                    .ToList();
            }
        }
        catch { }
    }

    private void SaveRecentFiles()
    {
        try
        {
            var dir = Path.GetDirectoryName(RecentFilesPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            File.WriteAllLines(RecentFilesPath, recentFiles.Take(MaxRecentFiles));
        }
        catch { }
    }

    private void AddToRecentFiles(string outputPath)
    {
        recentFiles.Remove(outputPath);
        recentFiles.Insert(0, outputPath);
        if (recentFiles.Count > MaxRecentFiles)
        {
            recentFiles = recentFiles.Take(MaxRecentFiles).ToList();
        }
        SaveRecentFiles();
    }
}
