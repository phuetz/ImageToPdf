using System.Runtime.InteropServices;

namespace ImageToPdf;

/// <summary>
/// Conversion par lot : l'utilisateur choisit plusieurs dossiers ; l'application
/// crée un PDF par dossier, regroupant toutes les images qu'il contient. Un écran
/// de progression (barre globale + barre du dossier courant + journal par dossier)
/// permet de suivre l'avancement et d'annuler à tout moment.
/// </summary>
public class BatchConvertForm : Form
{
    // Ordre « façon Explorateur Windows » : page2 avant page10 (tri naturel).
    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
    private static extern int StrCmpLogicalW(string a, string b);

    private static readonly IComparer<string> NaturalComparer =
        Comparer<string>.Create((a, b) => StrCmpLogicalW(a, b));

    // Une ligne de la liste : dossier source, nombre d'images, PDF produit (après run).
    private sealed class FolderRow
    {
        public string Folder { get; }
        public int Count { get; set; }
        public string? OutputPath { get; set; }
        public FolderRow(string folder, int count) { Folder = folder; Count = count; }
    }

    // Chemins déjà présents (déduplication rapide, insensible à la casse).
    private readonly HashSet<string> _folders = new(StringComparer.OrdinalIgnoreCase);

    private ListView _list = null!;
    private Button _btnAdd = null!;
    private Button _btnAddSub = null!;
    private Button _btnRemove = null!;
    private Button _btnClearAll = null!;
    private CheckBox _chkRecursive = null!;
    private RadioButton _rbInSource = null!;
    private RadioButton _rbInDest = null!;
    private TextBox _txtDestination = null!;
    private Button _btnBrowseDest = null!;
    private Label _lblOverall = null!;
    private ProgressBar _barOverall = null!;
    private Label _lblCurrent = null!;
    private ProgressBar _barCurrent = null!;
    private Button _btnStart = null!;
    private Button _btnOpenDest = null!;
    private Button _btnClose = null!;
    private ToolTip _tips = null!;

    private CancellationTokenSource? _cts;
    private bool _running;
    // Dossier à ouvrir via « Ouvrir le dossier » : la destination, ou le dernier
    // dossier source ayant reçu un PDF (mode « à côté des images »).
    private string? _openTargetDir;

    private static readonly Color OkColor = Color.FromArgb(0, 120, 60);
    private static readonly Color ErrColor = Color.FromArgb(180, 40, 30);
    private static readonly Color SkipColor = Color.Gray;

    public BatchConvertForm()
    {
        BuildUi();
    }

    private void BuildUi()
    {
        Text = "Conversion par lot";
        Size = new Size(720, 640);
        MinimumSize = new Size(620, 560);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 9);
        BackColor = Color.White;
        Icon = TryGetOwnerIcon();
        KeyPreview = true;
        AllowDrop = true;
        _tips = new ToolTip();

        // ── En-tête ────────────────────────────────────────────────────────────
        var lblHeader = new Label
        {
            Text = "Ajoutez des dossiers (bouton ou glisser-déposer) : un PDF sera créé "
                 + "pour chacun, regroupant toutes ses images.",
            Dock = DockStyle.Top,
            Padding = new Padding(14, 12, 14, 6),
            Height = 44,
            ForeColor = Color.FromArgb(60, 60, 60)
        };

        // ── Liste des dossiers ─────────────────────────────────────────────────
        _list = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            HideSelection = false,
            MultiSelect = true,
            AllowDrop = true,
            ShowItemToolTips = true
        };
        _list.Columns.Add("Dossier", 320);
        _list.Columns.Add("Images", 70, HorizontalAlignment.Right);
        _list.Columns.Add("État", 230);
        _list.Resize += (_, _) => AdjustColumns();
        _list.DoubleClick += (_, _) => OpenSelectedRow();
        _list.KeyDown += (_, e) => { if (e.KeyCode == Keys.Delete) { RemoveSelected(); e.Handled = true; } };
        _list.DragEnter += HandleDragEnter;
        _list.DragDrop += HandleDragDrop;

        var listPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(14, 4, 14, 4) };
        listPanel.Controls.Add(_list);

        // ── Boutons de gestion de la liste ─────────────────────────────────────
        _btnAdd = MakeButton("Ajouter un dossier…", 160);
        _btnAdd.Click += (_, _) => AddSingleFolder();

        _btnAddSub = MakeButton("Ajouter les sous-dossiers…", 190);
        _btnAddSub.Click += (_, _) => AddSubfolders();
        _tips.SetToolTip(_btnAddSub, "Choisir un dossier parent : chacun de ses sous-dossiers deviendra un PDF distinct.");

        _btnRemove = MakeButton("Retirer", 80);
        _btnRemove.Click += (_, _) => RemoveSelected();
        _tips.SetToolTip(_btnRemove, "Retirer les dossiers sélectionnés (touche Suppr).");

        _btnClearAll = MakeButton("Tout effacer", 100);
        _btnClearAll.Click += (_, _) => ClearAll();

        var listButtons = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(14, 4, 14, 4),
            Height = 44,
            AutoSize = false
        };
        listButtons.Controls.AddRange(new Control[] { _btnAdd, _btnAddSub, _btnRemove, _btnClearAll });

        // ── Options ────────────────────────────────────────────────────────────
        // TableLayoutPanel : gère le dimensionnement quel que soit l'ordre de
        // création (l'ancrage manuel dans un Panel se calcule sur la taille du
        // panneau à l'ajout — piège qui poussait le bouton « Parcourir » hors champ).
        _chkRecursive = new CheckBox
        {
            Text = "Inclure les images des sous-dossiers",
            AutoSize = true,
            Margin = new Padding(0, 6, 0, 6)
        };
        _chkRecursive.CheckedChanged += (_, _) => { RefreshCounts(); UpdateReadiness(); };

        // Où enregistrer les PDF : dans chaque dossier source, ou dans un dossier unique.
        _rbInSource = new RadioButton
        {
            Text = "Enregistrer chaque PDF dans son dossier d'images",
            AutoSize = true,
            Checked = true,
            Margin = new Padding(0, 6, 0, 0)
        };
        _rbInDest = new RadioButton
        {
            Text = "Enregistrer dans un dossier de destination :",
            AutoSize = true,
            Margin = new Padding(0, 4, 0, 0)
        };
        _rbInSource.CheckedChanged += (_, _) => { UpdateDestinationEnabled(); UpdateReadiness(); };

        _txtDestination = new TextBox
        {
            Dock = DockStyle.Fill,
            Margin = new Padding(24, 4, 8, 0)
        };
        _txtDestination.TextChanged += (_, _) => UpdateReadiness();
        _btnBrowseDest = MakeButton("Parcourir…", 100);
        _btnBrowseDest.Margin = new Padding(0, 2, 0, 0);
        _btnBrowseDest.Click += (_, _) => BrowseDestination();

        var destRow = new TableLayoutPanel
        {
            ColumnCount = 2,
            RowCount = 1,
            AutoSize = true,
            Dock = DockStyle.Top,
            Margin = new Padding(0)
        };
        destRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        destRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        destRow.Controls.Add(_txtDestination, 0, 0);
        destRow.Controls.Add(_btnBrowseDest, 1, 0);

        var optionsPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            RowCount = 4,
            AutoSize = true,
            Padding = new Padding(14, 6, 14, 6)
        };
        optionsPanel.Controls.Add(_chkRecursive, 0, 0);
        optionsPanel.Controls.Add(_rbInSource, 0, 1);
        optionsPanel.Controls.Add(_rbInDest, 0, 2);
        optionsPanel.Controls.Add(destRow, 0, 3);

        // ── Écran de progression ───────────────────────────────────────────────
        _lblOverall = new Label
        {
            Text = "Prêt.",
            AutoSize = false,
            Dock = DockStyle.Top,
            Height = 20,
            Padding = new Padding(14, 0, 14, 0)
        };
        _barOverall = new ProgressBar { Dock = DockStyle.Top, Height = 20 };
        _lblCurrent = new Label
        {
            Text = "",
            AutoSize = false,
            Dock = DockStyle.Top,
            Height = 20,
            Padding = new Padding(14, 4, 14, 0),
            ForeColor = Color.FromArgb(90, 90, 90)
        };
        _barCurrent = new ProgressBar { Dock = DockStyle.Top, Height = 16 };

        // Les ProgressBar en Dock ne respectent pas Margin : on les enveloppe.
        var progressPanel = new Panel { Dock = DockStyle.Bottom, Height = 96, Padding = new Padding(14, 4, 14, 4) };
        // Ajout en ordre inverse car Dock=Top empile du haut vers le bas.
        progressPanel.Controls.Add(_barCurrent);
        progressPanel.Controls.Add(_lblCurrent);
        progressPanel.Controls.Add(_barOverall);
        progressPanel.Controls.Add(_lblOverall);

        // ── Barre de boutons du bas ────────────────────────────────────────────
        _btnStart = MakeButton("Démarrer", 130);
        _btnStart.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        _btnStart.ForeColor = Color.FromArgb(0, 95, 184);
        _btnStart.Click += BtnStart_Click;

        _btnOpenDest = MakeButton("Ouvrir le dossier", 150);
        _btnOpenDest.Enabled = false;
        _btnOpenDest.Click += (_, _) => OpenPath(_openTargetDir);

        _btnClose = MakeButton("Fermer", 100);
        _btnClose.Click += (_, _) => Close();

        var bottomButtons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(14, 8, 14, 10),
            Height = 52,
            AutoSize = false
        };
        bottomButtons.Controls.AddRange(new Control[] { _btnStart, _btnClose, _btnOpenDest });

        // Ordre d'ajout au formulaire : Fill en premier, puis les Dock extérieurs.
        Controls.Add(listPanel);
        Controls.Add(listButtons);
        Controls.Add(optionsPanel);
        Controls.Add(lblHeader);
        Controls.Add(progressPanel);
        Controls.Add(bottomButtons);

        // Entrée = Démarrer, Échap = Fermer (géré par le garde de fermeture pendant un run).
        AcceptButton = _btnStart;
        CancelButton = _btnClose;
        DragEnter += HandleDragEnter;
        DragDrop += HandleDragDrop;
        FormClosing += BatchConvertForm_FormClosing;

        UpdateDestinationEnabled();
        UpdateReadiness();
    }

    // Les contrôles de destination ne servent qu'en mode « dossier de destination ».
    private void UpdateDestinationEnabled()
    {
        bool useDest = _rbInDest.Checked;
        _txtDestination.Enabled = useDest && !_running;
        _btnBrowseDest.Enabled = useDest && !_running;
    }

    // Ajuste la colonne « Dossier » pour occuper la largeur restante.
    private void AdjustColumns()
    {
        if (_list.Columns.Count < 3) return;
        int rest = _list.Columns[1].Width + _list.Columns[2].Width;
        _list.Columns[0].Width = Math.Max(160, _list.ClientSize.Width - rest - 4);
    }

    // Bilan en direct : nombre de dossiers/images, état de préparation, activation
    // du bouton « Démarrer ». Ne s'exécute qu'au repos (pendant un run, le libellé
    // affiche l'avancement).
    private void UpdateReadiness()
    {
        if (_running) return;

        int folders = _list.Items.Count;
        int totalImages = 0, withImages = 0;
        foreach (ListViewItem it in _list.Items)
        {
            if (it.Tag is FolderRow fr)
            {
                totalImages += fr.Count;
                if (fr.Count > 0) withImages++;
            }
        }

        bool destOk = _rbInSource.Checked || !string.IsNullOrWhiteSpace(_txtDestination.Text);
        _btnStart.Enabled = withImages > 0 && destOk;
        _btnRemove.Enabled = _list.SelectedItems.Count > 0 || folders > 0;
        _btnClearAll.Enabled = folders > 0;

        if (folders == 0)
        {
            _lblOverall.Text = "Glissez-déposez des dossiers ici, ou cliquez sur « Ajouter un dossier… ».";
        }
        else
        {
            string msg = $"{folders} dossier(s) · {totalImages} image(s)";
            if (withImages == 0) msg += " — aucune image à convertir.";
            else if (!destOk) msg += " — indiquez un dossier de destination.";
            else msg += $" — {withImages} PDF à créer. Cliquez sur « Démarrer ».";
            _lblOverall.Text = msg;
        }
    }

    private static Button MakeButton(string text, int width) => new()
    {
        Text = text,
        Width = width,
        Height = 30,
        FlatStyle = FlatStyle.System,
        Margin = new Padding(0, 0, 8, 0),
        AutoSize = false
    };

    private Icon? TryGetOwnerIcon()
    {
        try { return Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
        catch { return null; }
    }

    // ── Gestion de la liste de dossiers ────────────────────────────────────────

    private void AddSingleFolder()
    {
        using var dlg = new FolderBrowserDialog
        {
            Description = "Choisissez un dossier d'images",
            UseDescriptionForTitle = true
        };
        if (dlg.ShowDialog(this) != DialogResult.OK)
            return;

        AddFolder(dlg.SelectedPath);
        AfterFoldersChanged();
    }

    private void AddSubfolders()
    {
        using var dlg = new FolderBrowserDialog
        {
            Description = "Choisissez un dossier parent — chacun de ses sous-dossiers deviendra un PDF",
            UseDescriptionForTitle = true
        };
        if (dlg.ShowDialog(this) != DialogResult.OK)
            return;

        string[] subs;
        try { subs = Directory.GetDirectories(dlg.SelectedPath); }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Impossible de lire ce dossier :\n{ex.Message}", "Erreur",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (subs.Length == 0)
        {
            MessageBox.Show(this, "Ce dossier ne contient aucun sous-dossier.", "Aucun sous-dossier",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        _list.BeginUpdate();
        foreach (var sub in subs.OrderBy(s => s, NaturalComparer))
            AddFolder(sub);
        _list.EndUpdate();
        AfterFoldersChanged();
    }

    private void AddFolder(string folder)
    {
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            return;

        var full = Path.GetFullPath(folder);
        if (!_folders.Add(full)) // déjà présent
            return;

        int count = CountImages(full);
        var row = new FolderRow(full, count);

        // Colonne 1 = nom du dossier (lisible) ; chemin complet en info-bulle.
        var name = Path.GetFileName(full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        if (string.IsNullOrEmpty(name)) name = full;

        var item = new ListViewItem(name) { Tag = row, ToolTipText = full };
        item.SubItems.Add(count.ToString());
        item.SubItems.Add(count == 0 ? "⚠ Aucune image" : "En attente");
        if (count == 0) item.ForeColor = SkipColor;
        _list.Items.Add(item);
    }

    // Actions communes après tout ajout de dossiers.
    private void AfterFoldersChanged()
    {
        PrefillDestination();
        AdjustColumns();
        UpdateReadiness();
    }

    private void PrefillDestination()
    {
        if (!string.IsNullOrWhiteSpace(_txtDestination.Text) || _folders.Count == 0)
            return;
        var first = _folders.First();
        _txtDestination.Text = Path.GetDirectoryName(first.TrimEnd(Path.DirectorySeparatorChar)) ?? first;
    }

    private void RemoveSelected()
    {
        if (_running) return;
        foreach (ListViewItem item in _list.SelectedItems)
        {
            if (item.Tag is FolderRow fr)
                _folders.Remove(fr.Folder);
            item.Remove();
        }
        UpdateReadiness();
    }

    private void ClearAll()
    {
        if (_running) return;
        _folders.Clear();
        _list.Items.Clear();
        UpdateReadiness();
    }

    private void RefreshCounts()
    {
        foreach (ListViewItem item in _list.Items)
        {
            if (item.Tag is not FolderRow fr) continue;
            fr.Count = CountImages(fr.Folder);
            item.SubItems[1].Text = fr.Count.ToString();
            // Ne toucher l'état/couleur que si aucune conversion n'a encore eu lieu.
            if (fr.OutputPath == null)
            {
                item.SubItems[2].Text = fr.Count == 0 ? "⚠ Aucune image" : "En attente";
                item.ForeColor = fr.Count == 0 ? SkipColor : _list.ForeColor;
            }
        }
    }

    private int CountImages(string folder)
    {
        try { return EnumerateImages(folder).Count; }
        catch { return 0; }
    }

    private List<string> EnumerateImages(string folder)
    {
        var option = _chkRecursive.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        return Directory.EnumerateFiles(folder, "*", option)
            .Where(PdfMerger.IsImageFile)
            .OrderBy(f => f, NaturalComparer)
            .ToList();
    }

    private void BrowseDestination()
    {
        using var dlg = new FolderBrowserDialog
        {
            Description = "Dossier où enregistrer les PDF",
            UseDescriptionForTitle = true,
            SelectedPath = Directory.Exists(_txtDestination.Text) ? _txtDestination.Text : ""
        };
        if (dlg.ShowDialog(this) == DialogResult.OK)
            _txtDestination.Text = dlg.SelectedPath;
    }

    // Double-clic : ouvre le PDF produit s'il existe, sinon le dossier source.
    private void OpenSelectedRow()
    {
        if (_running || _list.SelectedItems.Count == 0) return;
        if (_list.SelectedItems[0].Tag is not FolderRow fr) return;
        OpenPath(fr.OutputPath != null && File.Exists(fr.OutputPath) ? fr.OutputPath : fr.Folder);
    }

    // ── Glisser-déposer de dossiers ────────────────────────────────────────────

    private void HandleDragEnter(object? sender, DragEventArgs e)
    {
        e.Effect = (!_running && e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
            ? DragDropEffects.Copy
            : DragDropEffects.None;
    }

    private void HandleDragDrop(object? sender, DragEventArgs e)
    {
        if (_running || e.Data?.GetData(DataFormats.FileDrop) is not string[] paths)
            return;

        _list.BeginUpdate();
        foreach (var p in paths)
        {
            if (Directory.Exists(p))
                AddFolder(p);
            else if (File.Exists(p)) // un fichier lâché → on prend son dossier
            {
                var dir = Path.GetDirectoryName(p);
                if (dir != null) AddFolder(dir);
            }
        }
        _list.EndUpdate();
        AfterFoldersChanged();
    }

    // ── Conversion ─────────────────────────────────────────────────────────────

    private async void BtnStart_Click(object? sender, EventArgs e)
    {
        // Pendant l'exécution, le bouton « Démarrer » devient « Annuler ».
        if (_running)
        {
            _cts?.Cancel();
            _btnStart.Enabled = false;
            _btnStart.Text = "Annulation…";
            return;
        }

        if (_folders.Count == 0)
        {
            MessageBox.Show(this, "Ajoutez au moins un dossier.", "Aucun dossier",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Deux modes : PDF à côté des images (dans le dossier source), ou dans un
        // dossier de destination unique.
        bool inSource = _rbInSource.Checked;
        var destination = _txtDestination.Text.Trim();

        if (!inSource)
        {
            if (string.IsNullOrEmpty(destination))
            {
                MessageBox.Show(this, "Choisissez un dossier de destination pour les PDF.", "Destination requise",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Directory.CreateDirectory(destination);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Impossible de créer le dossier de destination :\n{ex.Message}",
                    "Destination invalide", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        // Séquence des dossiers, alignée sur l'ordre d'affichage des lignes.
        var rows = new List<(int Index, FolderRow Row)>();
        for (int i = 0; i < _list.Items.Count; i++)
            if (_list.Items[i].Tag is FolderRow fr)
                rows.Add((i, fr));

        bool recursive = _chkRecursive.Checked;
        _openTargetDir = null;

        SetRunning(true);
        _cts = new CancellationTokenSource();
        var token = _cts.Token;

        _barOverall.Minimum = 0;
        _barOverall.Maximum = rows.Count;
        _barOverall.Value = 0;
        _barCurrent.Value = 0;

        int okCount = 0, skipCount = 0, errCount = 0;

        try
        {
            await Task.Run(() =>
            {
                // Clé de désambiguïsation = chemin de sortie complet (deux dossiers
                // « Scans » situés ailleurs ne se gênent donc pas entre eux).
                var usedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int d = 0; d < rows.Count; d++)
                {
                    token.ThrowIfCancellationRequested();

                    var (index, row) = rows[d];
                    var folder = row.Folder;
                    Ui(() =>
                    {
                        _lblOverall.Text = $"Dossier {d + 1} / {rows.Count} : {Path.GetFileName(folder)}";
                        SetRowState(index, "En cours…", _list.ForeColor);
                    });

                    List<string> images;
                    try
                    {
                        var option = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                        images = Directory.EnumerateFiles(folder, "*", option)
                            .Where(PdfMerger.IsImageFile)
                            .OrderBy(f => f, NaturalComparer)
                            .ToList();
                    }
                    catch (Exception ex)
                    {
                        errCount++;
                        Ui(() => { SetRowState(index, "✗ Erreur : " + ex.Message, ErrColor); _barOverall.Value = d + 1; });
                        continue;
                    }

                    if (images.Count == 0)
                    {
                        skipCount++;
                        Ui(() => { SetRowState(index, "⚠ Aucune image", SkipColor); _barOverall.Value = d + 1; });
                        continue;
                    }

                    var targetDir = inSource ? folder : destination;
                    var outPath = BuildOutputPath(targetDir, folder, usedPaths);

                    Ui(() =>
                    {
                        _barCurrent.Minimum = 0;
                        _barCurrent.Maximum = images.Count;
                        _barCurrent.Value = 0;
                    });

                    try
                    {
                        PdfMerger.CreatePdf(images, outPath, (cur, total, fileName) =>
                        {
                            Ui(() =>
                            {
                                _barCurrent.Maximum = Math.Max(total, 1);
                                _barCurrent.Value = Math.Min(cur, _barCurrent.Maximum);
                                _lblCurrent.Text = $"Image {cur} / {total} : {fileName}";
                            });
                        }, token);

                        okCount++;
                        _openTargetDir = targetDir; // dernier dossier ayant reçu un PDF
                        Ui(() =>
                        {
                            row.OutputPath = outPath;
                            SetRowState(index, $"✓ Créé — {Path.GetFileName(outPath)}", OkColor);
                        });
                    }
                    catch (OperationCanceledException)
                    {
                        // Le PDF partiel éventuel est supprimé, puis on remonte l'annulation.
                        TryDelete(outPath);
                        Ui(() => SetRowState(index, "Annulé", SkipColor));
                        throw;
                    }
                    catch (Exception ex)
                    {
                        errCount++;
                        TryDelete(outPath);
                        Ui(() => SetRowState(index, "✗ Erreur : " + ex.Message, ErrColor));
                    }

                    Ui(() => _barOverall.Value = d + 1);
                }
            }, token);

            _lblOverall.Text = $"Terminé : {okCount} PDF créé(s)"
                + (skipCount > 0 ? $", {skipCount} ignoré(s)" : "")
                + (errCount > 0 ? $", {errCount} en erreur" : "")
                + ".  Double-cliquez une ligne pour ouvrir son PDF.";
            _lblCurrent.Text = "";
            _barCurrent.Value = 0;
        }
        catch (OperationCanceledException)
        {
            _lblOverall.Text = $"Conversion annulée — {okCount} PDF créé(s) avant l'arrêt.";
            _lblCurrent.Text = "";
            _barCurrent.Value = 0;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Erreur inattendue :\n{ex.Message}", "Erreur",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _cts?.Dispose();
            _cts = null;
            SetRunning(false);
            _btnOpenDest.Enabled = okCount > 0 && _openTargetDir != null && Directory.Exists(_openTargetDir);
        }
    }

    // Nom du PDF = nom du dossier ; désambiguïsation si collision (dans le lot ou sur
    // disque). En mode « dans le dossier source », targetDir = le dossier lui-même.
    private static string BuildOutputPath(string targetDir, string folder, HashSet<string> usedPaths)
    {
        var baseName = SanitizeFileName(Path.GetFileName(folder.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)));
        if (string.IsNullOrWhiteSpace(baseName))
            baseName = "PDF";

        var path = Path.Combine(targetDir, baseName + ".pdf");
        int n = 2;
        while (usedPaths.Contains(path) || File.Exists(path))
        {
            path = Path.Combine(targetDir, $"{baseName} ({n}).pdf");
            n++;
        }
        usedPaths.Add(path);
        return path;
    }

    private static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name.Trim();
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* best effort */ }
    }

    private void SetRowState(int index, string state, Color color)
    {
        if (index < 0 || index >= _list.Items.Count) return;
        var item = _list.Items[index];
        item.SubItems[2].Text = state;
        item.ForeColor = color;
        item.EnsureVisible();
    }

    // Marshalle une action vers le thread UI (le worker tourne sur un thread de pool).
    private void Ui(Action action)
    {
        if (IsDisposed || Disposing) return;
        try
        {
            if (InvokeRequired) Invoke(action);
            else action();
        }
        catch (ObjectDisposedException) { /* formulaire fermé pendant la conversion */ }
        catch (InvalidOperationException) { /* handle détruit */ }
    }

    private void SetRunning(bool running)
    {
        _running = running;
        _btnAdd.Enabled = !running;
        _btnAddSub.Enabled = !running;
        _btnRemove.Enabled = !running;
        _btnClearAll.Enabled = !running;
        _chkRecursive.Enabled = !running;
        _rbInSource.Enabled = !running;
        _rbInDest.Enabled = !running;
        _list.Enabled = !running;
        UpdateDestinationEnabled(); // tient compte à la fois du mode et de _running
        _btnClose.Enabled = !running;

        _btnStart.Enabled = true;
        _btnStart.Text = running ? "Annuler" : "Démarrer";
        _btnStart.ForeColor = running ? Color.FromArgb(196, 43, 28) : Color.FromArgb(0, 95, 184);

        if (running) _btnOpenDest.Enabled = false;
    }

    private void OpenPath(string? path)
    {
        if (string.IsNullOrEmpty(path) || !(File.Exists(path) || Directory.Exists(path)))
            return;
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
        }
        catch { /* ignore */ }
    }

    private void BatchConvertForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        // Empêche la fermeture pendant une conversion : il faut d'abord annuler.
        if (_running)
        {
            var result = MessageBox.Show(this,
                "Une conversion est en cours. Voulez-vous l'annuler et fermer ?",
                "Conversion en cours", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
                _cts?.Cancel();
            e.Cancel = true; // on ne ferme pas maintenant ; l'utilisateur refermera après l'arrêt
        }
    }
}
