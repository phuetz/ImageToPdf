namespace ImageToPdf;

public class SettingsForm : Form
{
    private CheckBox chkPageNumbers = null!;
    private ComboBox cmbPageNumberPosition = null!;
    private CheckBox chkWatermark = null!;
    private TextBox txtWatermark = null!;
    private TrackBar trackWatermarkOpacity = null!;
    private Label lblOpacityValue = null!;
    private ComboBox cmbPageFormat = null!;
    private TrackBar trackImageQuality = null!;
    private Label lblQualityValue = null!;
    private Button btnOk = null!;
    private Button btnCancel = null!;

    public bool PageNumbersEnabled => chkPageNumbers.Checked;
    public int PageNumberPosition => cmbPageNumberPosition.SelectedIndex;
    public bool WatermarkEnabled => chkWatermark.Checked;
    public string WatermarkText => txtWatermark.Text;
    public int WatermarkOpacity => trackWatermarkOpacity.Value;
    public int PageFormat => cmbPageFormat.SelectedIndex;
    public int ImageQuality => trackImageQuality.Value;

    public SettingsForm(bool pageNumbers, int pageNumPos, bool watermark, string watermarkText,
        int watermarkOpacity, int pageFormat, int imageQuality)
    {
        InitializeComponent();

        chkPageNumbers.Checked = pageNumbers;
        cmbPageNumberPosition.SelectedIndex = pageNumPos;
        chkWatermark.Checked = watermark;
        txtWatermark.Text = watermarkText;
        trackWatermarkOpacity.Value = watermarkOpacity;
        cmbPageFormat.SelectedIndex = pageFormat;
        trackImageQuality.Value = imageQuality;

        UpdateWatermarkControls();
        UpdatePageNumberControls();
        UpdateOpacityLabel();
        UpdateQualityLabel();
    }

    private void InitializeComponent()
    {
        this.Text = "Paramètres du PDF - Pro";
        this.Size = new Size(450, 480);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.StartPosition = FormStartPosition.CenterParent;

        var mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(20)
        };

        int y = 10;

        // === Section Format de page ===
        var lblFormatSection = new Label
        {
            Text = "Format de page",
            Font = new Font(this.Font.FontFamily, 10, FontStyle.Bold),
            Location = new Point(10, y),
            AutoSize = true
        };
        y += 25;

        var lblFormat = new Label
        {
            Text = "Format pour les images :",
            Location = new Point(20, y + 3),
            AutoSize = true
        };

        cmbPageFormat = new ComboBox
        {
            Location = new Point(180, y),
            Width = 200,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        cmbPageFormat.Items.AddRange(new[] { "Adapter à l'image", "A4", "Letter", "A3" });
        y += 35;

        // === Section Qualité image ===
        var lblQualitySection = new Label
        {
            Text = "Qualité des images",
            Font = new Font(this.Font.FontFamily, 10, FontStyle.Bold),
            Location = new Point(10, y),
            AutoSize = true
        };
        y += 25;

        var lblQuality = new Label
        {
            Text = "Qualité (compression) :",
            Location = new Point(20, y + 3),
            AutoSize = true
        };

        trackImageQuality = new TrackBar
        {
            Location = new Point(180, y - 5),
            Width = 170,
            Minimum = 10,
            Maximum = 100,
            TickFrequency = 10,
            Value = 85
        };
        trackImageQuality.ValueChanged += (s, e) => UpdateQualityLabel();

        lblQualityValue = new Label
        {
            Text = "85%",
            Location = new Point(360, y + 3),
            AutoSize = true
        };
        y += 50;

        // === Section Numérotation ===
        var lblNumberSection = new Label
        {
            Text = "Numérotation des pages",
            Font = new Font(this.Font.FontFamily, 10, FontStyle.Bold),
            Location = new Point(10, y),
            AutoSize = true
        };
        y += 25;

        chkPageNumbers = new CheckBox
        {
            Text = "Ajouter les numéros de page",
            Location = new Point(20, y),
            AutoSize = true
        };
        chkPageNumbers.CheckedChanged += (s, e) => UpdatePageNumberControls();
        y += 28;

        var lblPosition = new Label
        {
            Text = "Position :",
            Location = new Point(40, y + 3),
            AutoSize = true
        };

        cmbPageNumberPosition = new ComboBox
        {
            Location = new Point(180, y),
            Width = 200,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        cmbPageNumberPosition.Items.AddRange(new[] { "Bas - Centre", "Bas - Droite", "Haut - Centre", "Haut - Droite" });
        y += 40;

        // === Section Filigrane ===
        var lblWatermarkSection = new Label
        {
            Text = "Filigrane",
            Font = new Font(this.Font.FontFamily, 10, FontStyle.Bold),
            Location = new Point(10, y),
            AutoSize = true
        };
        y += 25;

        chkWatermark = new CheckBox
        {
            Text = "Ajouter un filigrane",
            Location = new Point(20, y),
            AutoSize = true
        };
        chkWatermark.CheckedChanged += (s, e) => UpdateWatermarkControls();
        y += 28;

        var lblWatermarkText = new Label
        {
            Text = "Texte :",
            Location = new Point(40, y + 3),
            AutoSize = true
        };

        txtWatermark = new TextBox
        {
            Location = new Point(180, y),
            Width = 200,
            Text = ""
        };
        y += 30;

        var lblOpacity = new Label
        {
            Text = "Opacité :",
            Location = new Point(40, y + 3),
            AutoSize = true
        };

        trackWatermarkOpacity = new TrackBar
        {
            Location = new Point(180, y - 5),
            Width = 170,
            Minimum = 10,
            Maximum = 100,
            TickFrequency = 10,
            Value = 30
        };
        trackWatermarkOpacity.ValueChanged += (s, e) => UpdateOpacityLabel();

        lblOpacityValue = new Label
        {
            Text = "30%",
            Location = new Point(360, y + 3),
            AutoSize = true
        };
        y += 55;

        // === Boutons ===
        btnOk = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Location = new Point(250, y),
            Size = new Size(80, 30)
        };

        btnCancel = new Button
        {
            Text = "Annuler",
            DialogResult = DialogResult.Cancel,
            Location = new Point(340, y),
            Size = new Size(80, 30)
        };

        mainPanel.Controls.AddRange(new Control[]
        {
            lblFormatSection, lblFormat, cmbPageFormat,
            lblQualitySection, lblQuality, trackImageQuality, lblQualityValue,
            lblNumberSection, chkPageNumbers, lblPosition, cmbPageNumberPosition,
            lblWatermarkSection, chkWatermark, lblWatermarkText, txtWatermark,
            lblOpacity, trackWatermarkOpacity, lblOpacityValue,
            btnOk, btnCancel
        });

        this.Controls.Add(mainPanel);
        this.AcceptButton = btnOk;
        this.CancelButton = btnCancel;
    }

    private void UpdatePageNumberControls()
    {
        cmbPageNumberPosition.Enabled = chkPageNumbers.Checked;
    }

    private void UpdateWatermarkControls()
    {
        txtWatermark.Enabled = chkWatermark.Checked;
        trackWatermarkOpacity.Enabled = chkWatermark.Checked;
    }

    private void UpdateOpacityLabel()
    {
        lblOpacityValue.Text = $"{trackWatermarkOpacity.Value}%";
    }

    private void UpdateQualityLabel()
    {
        lblQualityValue.Text = $"{trackImageQuality.Value}%";
    }
}
