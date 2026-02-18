namespace CloudAdmin365.UI;

using System.Windows.Forms;
using CloudAdmin365.Services;
using CloudAdmin365.Utilities;

/// <summary>
/// Login dialog for user authentication.
/// Shows login status and connection information.
/// </summary>
public partial class LoginDialog : Form
{
    private readonly IAuthService _authService;
    private bool _loginSuccessful;

    public bool LoginSuccessful => _loginSuccessful;

    public LoginDialog(IAuthService authService)
    {
        ArgumentNullException.ThrowIfNull(authService);
        _authService = authService;
        InitializeComponent();
        SetupTheme();
    }

    private void InitializeComponent()
    {
        Text = "CloudAdmin365 — Login";
        Size = new Size(450, 320);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.White;

        Icon = IconGenerator.GetAppIcon();

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(20)
        };

        // Row 0: Logo / title
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        var lblTitle = new Label
        {
            Text = "CloudAdmin365",
            Font = AppTheme.TitleFont,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = AppTheme.ThemeBlue
        };
        layout.Controls.Add(lblTitle, 0, 0);

        // Row 1: Status message
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 65));
        var lblStatus = new Label
        {
            Text = "Click below to login with your Office 365 account.",
            Dock = DockStyle.Fill,
            AutoSize = false,
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = AppTheme.ThemeDimGray,
            Font = AppTheme.DefaultFont
        };
        layout.Controls.Add(lblStatus, 0, 1);

        // Row 2: Login button
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        var btnLogin = new Button
        {
            Text = "Login with Office 365",
            Font = AppTheme.BoldFont,
            BackColor = AppTheme.ThemeGreen,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Dock = DockStyle.Fill,
            Height = 45
        };
        btnLogin.Click += async (s, e) => await LoginButton_Click();
        layout.Controls.Add(btnLogin, 0, 2);

        // Row 3: Footer info
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var lblFooter = new Label
        {
            Text = "You will be prompted to authenticate in your browser.",
            Dock = DockStyle.Fill,
            AutoSize = false,
            TextAlign = ContentAlignment.TopCenter,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        layout.Controls.Add(lblFooter, 0, 3);

        Controls.Add(layout);
    }

    private void SetupTheme()
    {
        Font = AppTheme.DefaultFont;
    }

    private async Task LoginButton_Click()
    {
        try
        {
            var result = await _authService.LoginAsync();

            if (result.Success && result.User != null)
            {
                _loginSuccessful = true;
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show(
                    $"Login failed: {result.Error}",
                    "Login Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Login error: {ex.Message}",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
