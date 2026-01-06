using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows.Forms;

namespace QuickOptionImportExport;

public sealed class MainForm : Form
{
    private readonly DataGridView _grid = new();
    private readonly BindingList<OptionRow> _rows = new();
    private readonly BindingSource _bindingSource = new();

    private readonly Label _loadedLabel = new() { AutoSize = true, Text = "Loaded: 0 options" };
    private readonly TextBox _fileNameText = new() { ReadOnly = false };
    private readonly TextBox _fileLocationText = new() { ReadOnly = true };

    private readonly Button _browseLocationButton = new() { Text = "...", AutoSize = true };

    private readonly Button _importButton = new() { Text = "Import", AutoSize = true };
    private readonly Button _saveAsButton = new() { Text = "Save As", AutoSize = true };
    private readonly Button _saveButton = new() { Text = "Save", AutoSize = true };

    private readonly ContextMenuStrip _rowMenu = new();

    private string? _currentFilePath;
    private bool _dirty;

    public MainForm()
    {
        Text = "Quick Option Import/Export";
        StartPosition = FormStartPosition.CenterScreen;
        Width = 1180;
        Height = 760;
        MinimumSize = new System.Drawing.Size(980, 620);

        BuildUi();
        WireEvents();

        NewFile();
    }

    private void BuildUi()
    {
        // Grid
        _bindingSource.DataSource = _rows;

        _grid.Dock = DockStyle.Fill;
        _grid.DataSource = _bindingSource;
        _grid.AllowUserToAddRows = true;
        _grid.AllowUserToDeleteRows = false;
        _grid.AllowUserToResizeRows = false;
        _grid.RowHeadersVisible = true;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.MultiSelect = true;
        _grid.AutoGenerateColumns = false;
        _grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
        _grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

        _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        _grid.Columns.Add(MakeTextCol("OptionNumber", "Option Number", minWidth: 110));
        _grid.Columns.Add(MakeTextCol("Sequence", "Sequence", minWidth: 80));
        _grid.Columns.Add(MakeTextCol("TextValue", "Text Value", minWidth: 150));
        _grid.Columns.Add(MakeTextCol("TimeValue", "Time Value", minWidth: 110));
        _grid.Columns.Add(MakeTextCol("NumericValue", "Numeric Value", minWidth: 110, rightAlign: true));
        _grid.Columns.Add(MakeTextCol("LongValue", "Long Value", minWidth: 110, rightAlign: true));
        _grid.Columns.Add(MakeTextCol("DateValue", "Date Value", minWidth: 120));
        _grid.Columns.Add(new DataGridViewCheckBoxColumn
        {
            DataPropertyName = "BooleanValue",
            HeaderText = "Boolean Value",
            MinimumWidth = 110,
            FillWeight = 80,
            ThreeState = false
        });
        _grid.Columns.Add(MakeTextCol("AsciiValue", "Ascii Value", minWidth: 110));

        // Bottom bar (TableLayoutPanel keeps alignment clean and resizes well)
        var bottom = new TableLayoutPanel
        {
            Dock = DockStyle.Bottom,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(10, 8, 10, 10),
            ColumnCount = 7,
            RowCount = 2
        };

        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // Loaded
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // File Name label
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));  // File Name textbox (fill)
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110)); // Import
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110)); // Save As
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110)); // Save
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // Browse folder (row 2)

        bottom.RowStyles.Clear();
        bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
        bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var fileNameLabel = new Label { AutoSize = true, Text = "File Name", Anchor = AnchorStyles.Left };
        var fileLocationLabel = new Label { AutoSize = true, Text = "File Location", Anchor = AnchorStyles.Left };

        _fileNameText.Dock = DockStyle.Fill;
        _fileLocationText.Dock = DockStyle.Fill;

        fileLocationLabel.Margin = new Padding(0, 6, 6, 6);
        _fileLocationText.Margin = new Padding(0, 4, 0, 4);

        fileLocationLabel.Margin = new Padding(0, 10, 6, 0);
        _fileLocationText.Margin = new Padding(0, 7, 0, 0);

        bottom.Controls.Add(_loadedLabel, 0, 0);
        bottom.Controls.Add(fileNameLabel, 1, 0);
        bottom.Controls.Add(_fileNameText, 2, 0);
        bottom.Controls.Add(_importButton, 3, 0);
        bottom.Controls.Add(_saveAsButton, 4, 0);
        bottom.Controls.Add(_saveButton, 5, 0);

        bottom.Controls.Add(fileLocationLabel, 0, 1);
        bottom.Controls.Add(_fileLocationText, 1, 1);
        bottom.SetColumnSpan(_fileLocationText, 5);
        bottom.Controls.Add(_browseLocationButton, 6, 1);

        _browseLocationButton.Margin = new Padding(8, 4, 0, 4);
        _browseLocationButton.Anchor = AnchorStyles.Left;

        Controls.Add(_grid);
        Controls.Add(bottom);

        _importButton.Margin = new Padding(8, 4, 8, 4);
        _saveAsButton.Margin = new Padding(0, 4, 8, 4);
        _saveButton.Margin   = new Padding(0, 4, 0, 4);

        _importButton.Anchor = AnchorStyles.Left;
        _saveAsButton.Anchor = AnchorStyles.Left;
        _saveButton.Anchor   = AnchorStyles.Left;

        _importButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        _saveAsButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        _saveButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;


        // Context menu
        _rowMenu.Items.Add("Delete Row", null, (_, __) => DeleteSelectedRows());
    }

    private static DataGridViewTextBoxColumn MakeTextCol(string prop, string header, int minWidth, bool rightAlign = false)
    {
        return new DataGridViewTextBoxColumn
        {
            DataPropertyName = prop,
            HeaderText = header,
            MinimumWidth = minWidth,
            SortMode = DataGridViewColumnSortMode.Automatic,
            DefaultCellStyle = rightAlign
                ? new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight }
                : new DataGridViewCellStyle()
        };
    }

    private void WireEvents()
    {
        _importButton.Click += (_, __) => ImportJson();
        _saveButton.Click += (_, __) => Save();
        _saveAsButton.Click += (_, __) => SaveAs();
        _browseLocationButton.Click += (_, __) => BrowseForFolder();

        _grid.CellMouseDown += Grid_CellMouseDown;
        _grid.UserDeletingRow += (_, e) => { e.Cancel = true; };

        _grid.CellBeginEdit += (_, __) => MarkDirty();
        _grid.CellValueChanged += (_, __) => MarkDirty();
        _grid.UserAddedRow += (_, __) => { MarkDirty(); UpdateLoadedLabel(); };
        _grid.UserDeletedRow += (_, __) => { MarkDirty(); UpdateLoadedLabel(); };

        _bindingSource.ListChanged += (_, e) =>
        {
            if (e.ListChangedType == ListChangedType.ItemDeleted || e.ListChangedType == ListChangedType.ItemAdded)
                UpdateLoadedLabel();
        };

        _grid.CellValidating += Grid_CellValidating;
        _grid.DataError += (_, e) =>
        {
            MessageBox.Show(this, e.Exception.Message, "Invalid value", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            e.ThrowException = false;
        };

        FormClosing += MainForm_FormClosing;

        // Drag & drop a JSON file onto the grid
        AllowDrop = true;
        DragEnter += (_, e) =>
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true) e.Effect = DragDropEffects.Copy;
        };
        DragDrop += (_, e) =>
        {
            var files = (string[]?)e.Data?.GetData(DataFormats.FileDrop);
            var first = files?.FirstOrDefault(f => f.EndsWith(".json", StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(first)) ImportJson(first);
        };

        // Keep Save button state accurate
        UpdateTitle();
    }

    private void Grid_CellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button != MouseButtons.Right) return;
        if (e.RowIndex < 0) return;

        _grid.ClearSelection();
        _grid.Rows[e.RowIndex].Selected = true;
        _grid.CurrentCell = _grid.Rows[e.RowIndex].Cells[Math.Max(0, e.ColumnIndex)];
        _rowMenu.Show(Cursor.Position);
    }

    private void Grid_CellValidating(object? sender, DataGridViewCellValidatingEventArgs e)
    {
        if (e.RowIndex < 0) return;
        if (_grid.Rows[e.RowIndex].IsNewRow) return;

        var col = _grid.Columns[e.ColumnIndex];
        var header = col.HeaderText;
        var input = (e.FormattedValue?.ToString() ?? "").Trim();

        try
        {
            if (col.DataPropertyName == "OptionNumber" || col.DataPropertyName == "Sequence")
            {
                if (!int.TryParse(input, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                    throw new FormatException("Enter a whole number.");
            }
            else if (col.DataPropertyName == "NumericValue")
            {
                if (!double.TryParse(input, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                    throw new FormatException("Enter a numeric value (example: 0 or 12.34).");
            }
            else if (col.DataPropertyName == "LongValue")
            {
                if (!long.TryParse(input, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                    throw new FormatException("Enter a whole number.");
            }
            else if (col.DataPropertyName == "DateValue")
            {
                if (string.IsNullOrWhiteSpace(input)) return;
                _ = DateOnly.ParseExact(input, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            else if (col.DataPropertyName == "TimeValue")
            {
                if (string.IsNullOrWhiteSpace(input)) return;
                _ = OptionRow.ParseTime(input);
            }
        }
        catch (Exception ex)
        {
            e.Cancel = true;
            MessageBox.Show(this, $"{header} is invalid.\n\n{ex.Message}\n\nExpected:\n- Date: yyyy-MM-dd\n- Time: HH:mm or HH:mm:ss", "Invalid value", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (!_dirty) return;

        var result = MessageBox.Show(
            this,
            "You have unsaved changes. Save before closing?",
            "Unsaved Changes",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Warning);

        if (result == DialogResult.Cancel)
        {
            e.Cancel = true;
            return;
        }

        if (result == DialogResult.Yes)
        {
            if (!Save())
                e.Cancel = true;
        }
    }

    private void MarkDirty()
    {
        if (!_dirty)
        {
            _dirty = true;
            UpdateTitle();
        }
    }

    private void ClearDirty()
    {
        _dirty = false;
        UpdateTitle();
    }

    private void UpdateTitle()
    {
        var file = string.IsNullOrWhiteSpace(_currentFilePath) ? "(new file)" : Path.GetFileName(_currentFilePath);
        Text = _dirty ? $"Quick Option Import/Export - {file} *" : $"Quick Option Import/Export - {file}";
    }

    private void UpdateLoadedLabel() => _loadedLabel.Text = $"Loaded: {_rows.Count} options";

    private void NewFile()
    {
        _rows.Clear();
        _currentFilePath = null;
        _fileNameText.Text = "";
        _fileLocationText.Text = "";

        // Reset template to the minimal default template.
        OptionRow.TemplateFactory = OptionRow.CreateDefaultTemplate;

        UpdateLoadedLabel();
        ClearDirty();
    }

    private void ImportJson(string? filePath = null)
    {
        if (!ConfirmDiscardIfDirty()) return;

        filePath ??= PickJsonFileToOpen();
        if (string.IsNullOrWhiteSpace(filePath)) return;

        try
        {
            var text = File.ReadAllText(filePath, Encoding.UTF8);
            var node = JsonNode.Parse(text) ?? throw new InvalidDataException("File is empty or invalid JSON.");
            if (node is not JsonArray arr)
                throw new InvalidDataException("Expected the root JSON to be an array of option objects.");

            var objects = arr.OfType<JsonObject>().ToList();
            if (objects.Count != arr.Count)
                throw new InvalidDataException("One or more items in the JSON array is not an object.");

            // IMPORTANT: JsonNode instances retain their parent (the parsed JsonArray). If we
            // keep references to those nodes and then build a new JsonArray on save, System.Text.Json
            // will throw: "The node already has a parent.".
            // Fix: deep-clone every imported object so the grid edits operate on detached nodes.
            _rows.Clear();
            foreach (var obj in objects)
                _rows.Add(new OptionRow((JsonObject)obj.DeepClone()));

            // Set template to clone-first-row for perfect round-trip formatting.
            if (_rows.Count > 0)
            {
                var template = (JsonObject)_rows[0].Node.DeepClone();
                OptionRow.TemplateFactory = () => (JsonObject)template.DeepClone();
            }

            _currentFilePath = filePath;
            _fileNameText.Text = Path.GetFileName(filePath);
            _fileLocationText.Text = Path.GetDirectoryName(filePath) ?? "";

            UpdateLoadedLabel();
            ClearDirty();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Import Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private string? PickJsonFileToOpen()
    {
        using var ofd = new OpenFileDialog
        {
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            Title = "Import JSON options file"
        };

        return ofd.ShowDialog(this) == DialogResult.OK ? ofd.FileName : null;
    }

    private void BrowseForFolder()
    {
        using var fbd = new FolderBrowserDialog
        {
            Description = "Choose the folder where JSON files should be saved",
            UseDescriptionForTitle = true
        };

        if (Directory.Exists(_fileLocationText.Text))
            fbd.SelectedPath = _fileLocationText.Text;
        else if (!string.IsNullOrWhiteSpace(_currentFilePath))
            fbd.SelectedPath = Path.GetDirectoryName(_currentFilePath) ?? "";

        if (fbd.ShowDialog(this) == DialogResult.OK)
            _fileLocationText.Text = fbd.SelectedPath;
    }

    private bool Save()
    {
        if (string.IsNullOrWhiteSpace(_currentFilePath))
            return SaveAs();

        try
        {
            SaveToPath(_currentFilePath);
            _fileNameText.Text = Path.GetFileName(_currentFilePath);
            _fileLocationText.Text = Path.GetDirectoryName(_currentFilePath) ?? "";
            ClearDirty();
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Save Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }

    private bool SaveAs()
    {
        // Save As writes to the folder + file name shown in the bottom bar.
        // This matches the workflow of the legacy tool (type a new file name, click Save As).

        var dir = (_fileLocationText.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
        {
            BrowseForFolder();
            dir = (_fileLocationText.Text ?? "").Trim();
        }

        if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
        {
            MessageBox.Show(this,
                "Please choose a valid File Location folder.",
                "Save As",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return false;
        }

        var name = (_fileNameText.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show(this,
                "Please enter a File Name (example: TA-MyOptions.json).",
                "Save As",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            _fileNameText.Focus();
            return false;
        }

        if (!name.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            name += ".json";

        var targetPath = Path.Combine(dir, name);

        if (File.Exists(targetPath))
        {
            var overwrite = MessageBox.Show(this,
                "That file already exists. Overwrite it?",
                "Save As",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (overwrite != DialogResult.Yes)
                return false;
        }

        try
        {
            SaveToPath(targetPath);
            _currentFilePath = targetPath;
            _fileNameText.Text = Path.GetFileName(targetPath);
            _fileLocationText.Text = Path.GetDirectoryName(targetPath) ?? "";
            ClearDirty();
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Save Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }

    private void SaveToPath(string path)
    {
        // Commit pending edits before exporting.
        _grid.EndEdit();
        _bindingSource.EndEdit();

        var arr = new JsonArray();
        foreach (var r in _rows)
        {
            // Always deep-clone on write so the JSON we're serializing has no parent/child
            // conflicts and so we never mutate the in-memory objects during serialization.
            arr.Add(r.Node.DeepClone());
        }

        var json = arr.ToJsonString(new JsonSerializerOptions
        {
            WriteIndented = true
        });

        File.WriteAllText(path, json, Encoding.UTF8);
        UpdateTitle();
    }

    private void DeleteSelectedRows()
    {
        if (_grid.SelectedRows.Count == 0) return;

        var result = MessageBox.Show(this,
            $"Delete {_grid.SelectedRows.Count} selected row(s)?",
            "Delete Row",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (result != DialogResult.Yes) return;

        // Remove from end to start to preserve indices.
        var indices = _grid.SelectedRows
            .Cast<DataGridViewRow>()
            .Where(r => !r.IsNewRow)
            .Select(r => r.Index)
            .OrderByDescending(i => i)
            .ToList();

        foreach (var i in indices)
        {
            if (i >= 0 && i < _rows.Count)
                _rows.RemoveAt(i);
        }

        MarkDirty();
        UpdateLoadedLabel();
    }

    private bool ConfirmDiscardIfDirty()
    {
        if (!_dirty) return true;

        var result = MessageBox.Show(this,
            "You have unsaved changes. Discard them?",
            "Unsaved Changes",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        return result == DialogResult.Yes;
    }
}
