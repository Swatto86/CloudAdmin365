namespace CloudAdmin365.Utilities;

/// <summary>
/// Application theme colors.
/// </summary>
public static class AppTheme
{
    public static readonly Color ThemeBlue = Color.FromArgb(0, 120, 212);
    public static readonly Color ThemeGreen = Color.FromArgb(16, 124, 16);
    public static readonly Color ThemeRed = Color.FromArgb(200, 50, 50);
    public static readonly Color ThemeOrange = Color.FromArgb(230, 126, 34);
    public static readonly Color ThemeHeaderBg = Color.FromArgb(0, 120, 212);
    public static readonly Color ThemeHeaderFg = Color.White;
    public static readonly Color ThemeAltRow = Color.FromArgb(240, 246, 252);
    public static readonly Color ThemeDimGray = Color.DimGray;
    public static readonly Color ThemeBorder = Color.FromArgb(200, 200, 200);

    public static readonly Font DefaultFont = new("Segoe UI", 9);
    public static readonly Font BoldFont = new("Segoe UI", 9, FontStyle.Bold);
    public static readonly Font HeaderFont = new("Segoe UI", 9, FontStyle.Bold);
    public static readonly Font TitleFont = new("Segoe UI", 14, FontStyle.Bold);
}

/// <summary>
/// Common UI helpers.
/// </summary>
public static class UiHelpers
{
    public static DataGridView CreateThemedDataGrid(string[] columns, int[]? columnWeights = null)
    {
        var dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            RowHeadersVisible = false,
            EnableHeadersVisualStyles = false,
            ColumnHeadersHeight = 34
        };

        var headerStyle = new DataGridViewCellStyle
        {
            BackColor = AppTheme.ThemeHeaderBg,
            ForeColor = AppTheme.ThemeHeaderFg,
            Font = AppTheme.HeaderFont
        };
        dgv.ColumnHeadersDefaultCellStyle = headerStyle;

        dgv.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = AppTheme.ThemeAltRow
        };
        dgv.Font = AppTheme.DefaultFont;

        for (int i = 0; i < columns.Length; i++)
        {
            string colName = "col" + columns[i].Replace(" ", "");
            var column = new DataGridViewTextBoxColumn { Name = colName, HeaderText = columns[i] };
            if (columnWeights != null && i < columnWeights.Length)
                column.FillWeight = columnWeights[i];
            dgv.Columns.Add(column);
        }

        // Add context menu for copy
        var ctx = new ContextMenuStrip();
        var miCopyCell = new ToolStripMenuItem { Text = "Copy Cell" };
        miCopyCell.Click += (s, e) =>
        {
            var value = dgv.CurrentCell?.Value;
            if (value != null)
                Clipboard.SetText(value.ToString() ?? "");
        };
        var miCopyRow = new ToolStripMenuItem { Text = "Copy Row" };
        miCopyRow.Click += (s, e) =>
        {
            if (dgv.CurrentRow != null)
            {
                var values = dgv.CurrentRow.Cells.Cast<DataGridViewCell>()
                    .Select(c => c.Value?.ToString() ?? "")
                    .ToArray();
                Clipboard.SetText(string.Join("\t", values));
            }
        };
        ctx.Items.AddRange(new[] { miCopyCell, miCopyRow });
        dgv.ContextMenuStrip = ctx;

        return dgv;
    }

    public static void ExportDataGridToCsv(DataGridView grid, string defaultName)
    {
        var dlg = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
            DefaultExt = "csv",
            FileName = $"{defaultName}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dlg.ShowDialog() == DialogResult.OK)
        {
            try
            {
                using var writer = new StreamWriter(dlg.FileName, false, System.Text.Encoding.UTF8);
                
                // Write headers
                var headers = grid.Columns.Cast<DataGridViewColumn>()
                    .Select(col => $"\"{col.HeaderText}\"")
                    .ToArray();
                writer.WriteLine(string.Join(",", headers));

                // Write rows
                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        var values = row.Cells.Cast<DataGridViewCell>()
                            .Select(cell => $"\"{(cell.Value?.ToString() ?? "").Replace("\"", "\"\"")}\"")
                            .ToArray();
                        writer.WriteLine(string.Join(",", values));
                    }
                }

                MessageBox.Show($"Exported {grid.Rows.Count} row(s) to:\n{dlg.FileName}", 
                    "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed:\n{ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    public static void ShowStatus(Label label, string text, Color? color = null)
    {
        label.Text = text;
        label.ForeColor = color ?? AppTheme.ThemeDimGray;
        Application.DoEvents();
    }
}
