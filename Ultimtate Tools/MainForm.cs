#nullable enable
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using EtabsLoadAssigner.Core;

namespace EtabsLoadAssigner.Plugin;

/// <summary>
/// Cua so chinh cua plugin — port giao dien Tkinter 5 tab cua app Python
/// (etabs_loader/app.py) sang WinForms. UI build bang code (khong co file
/// .Designer) de de review tay.
///
/// 5 tab tuong ung 5 chuc nang:
///   1. File &amp; bang (tu dong nhan dien tu Excel)
///   2. Diem gan tai (+ alias diem Excel -> joint ETABS)
///   3. Phuong &amp; chieu Hx/Hy/Vz -> ETABS global
///   4. Match pattern (dong Excel &lt;-&gt; pattern ETABS)
///   5. Preview &amp; gan vao ETABS
///
/// Luu y: plugin chay trong tien trinh ETABS — KHONG goi
/// ApplicationConfiguration.Initialize() / EnableVisualStyles().
/// Cai dat luu ra %AppData%\EtabsLoadAssigner\profile.json (cPlugin bi huy
/// giua cac lan goi tu menu Tools nen khong giu state trong bo nho).
/// </summary>
public sealed class MainForm : Form
{
    private readonly IEtabsApi _liveApi;
    private readonly MockEtabs _mock = new();
    private IEtabsApi _etabs;

    private FoundationTable? _table;
    private Profile _profile = new();
    private bool _updatingSheet;

    private static string DefaultProfilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "EtabsLoadAssigner", "profile.json");

    // --- controls ---
    private readonly TextBox _fileTextBox = new() { Width = 380 };
    private readonly ComboBox _sheetCombo = new() { Width = 170, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly Label _statusLabel = new() { Dock = DockStyle.Top, Height = 26, ForeColor = Color.FromArgb(0x55, 0x55, 0x55) };

    private readonly DataGridView _grid = new()
    {
        Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false,
        RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
    };
    private readonly TextBox _fileInfoBox = new() { Dock = DockStyle.Bottom, Height = 110, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };

    private readonly ComboBox _pointCombo = new() { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly FlowLayoutPanel _aliasPanel = new()
    {
        Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown,
        WrapContents = false, AutoScroll = true,
    };
    private readonly Dictionary<string, TextBox> _aliasControls = new();
    private readonly TextBox _etabsInfoBox = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
    private readonly RadioButton _modeLiveRadio = new() { Text = "Model ETABS dang mo", Checked = true, AutoSize = true };
    private readonly RadioButton _modeMockRadio = new() { Text = "Offline — Mock ETABS (test, khong can ETABS)", AutoSize = true };

    private readonly Dictionary<string, (ComboBox Axis, ComboBox Sign)> _axisControls = new();
    private readonly Label _axisExampleLabel = new() { Dock = DockStyle.Top, AutoSize = true };

    private readonly FlowLayoutPanel _matchPanel = new()
    {
        Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown,
        WrapContents = false, AutoScroll = true,
    };
    private readonly Dictionary<string, ComboBox> _patternControls = new();
    private readonly CheckBox _createMissingCheck = new() { Text = "Tu tao pattern chua co trong model", AutoSize = true, Checked = true };
    private readonly ComboBox _reversibleCombo = new() { Width = 110, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox _rangeCombo = new() { Width = 110, DropDownStyle = ComboBoxStyle.DropDownList };

    private readonly CheckBox _replaceCheck = new() { Text = "Ghi de tai cu cung pattern tai diem nay", Checked = true, AutoSize = true };
    private readonly TextBox _previewBox = new()
    {
        Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, WordWrap = false,
        ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10),
    };

    public MainForm(IEtabsApi liveApi)
    {
        _liveApi = liveApi;
        _etabs = liveApi;
        Text = "ETABS Load Assigner v0.1 (Plugin) — Gan tai mong tu Excel vao ETABS";
        Size = new Size(1050, 720);
        StartPosition = FormStartPosition.CenterScreen;
        BuildUi();
        TryAutoLoadProfile();
        RefreshEtabsInfo();
    }

    // ---------------- UI ----------------
    private void BuildUi()
    {
        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 42, FlowDirection = FlowDirection.LeftToRight };
        var browseBtn = new Button { Text = "Chon file Excel...", AutoSize = true };
        browseBtn.Click += (_, _) => ChooseFile();
        top.Controls.Add(browseBtn);
        top.Controls.Add(_fileTextBox);
        top.Controls.Add(new Label { Text = "Sheet:", AutoSize = true });
        _sheetCombo.SelectedIndexChanged += (_, _) => { if (!_updatingSheet) Detect(); };
        top.Controls.Add(_sheetCombo);
        var detectBtn = new Button { Text = "Nhan dien bang", AutoSize = true };
        detectBtn.Click += (_, _) => Detect();
        top.Controls.Add(detectBtn);
        var loadBtn = new Button { Text = "Tai profile...", AutoSize = true };
        loadBtn.Click += (_, _) => LoadProfile();
        top.Controls.Add(loadBtn);
        var saveBtn = new Button { Text = "Luu profile...", AutoSize = true };
        saveBtn.Click += (_, _) => SaveProfile();
        top.Controls.Add(saveBtn);
        Controls.Add(top);
        _statusLabel.Text = "Chua import file.";
        Controls.Add(_statusLabel);

        var tabs = new TabControl { Dock = DockStyle.Fill };
        tabs.TabPages.Add(BuildTabFile());
        tabs.TabPages.Add(BuildTabPoint());
        tabs.TabPages.Add(BuildTabAxis());
        tabs.TabPages.Add(BuildTabMatch());
        tabs.TabPages.Add(BuildTabAssign());
        Controls.Add(tabs);
    }

    private TabPage BuildTabFile()
    {
        var page = new TabPage("1. File & bang");
        page.Controls.Add(_grid);
        page.Controls.Add(_fileInfoBox);
        page.Controls.Add(new Label
        {
            Dock = DockStyle.Top, AutoSize = true,
            Text = "Bang tai mong da nhan dien (chi doc — sua o file goc roi nhan dien lai):",
        });
        return page;
    }

    private TabPage BuildTabPoint()
    {
        var page = new TabPage("2. Diem gan tai");
        page.Controls.Add(_aliasPanel);
        var bottom = new Panel { Dock = DockStyle.Bottom, Height = 170 };
        bottom.Controls.Add(_etabsInfoBox);
        var modeBox = new GroupBox { Text = "ETABS", Dock = DockStyle.Top, Height = 70 };
        _modeLiveRadio.Location = new Point(10, 18);
        _modeLiveRadio.CheckedChanged += (_, _) => { if (_modeLiveRadio.Checked) SwitchMode(); };
        _modeMockRadio.Location = new Point(10, 42);
        _modeMockRadio.CheckedChanged += (_, _) => { if (_modeMockRadio.Checked) SwitchMode(); };
        modeBox.Controls.Add(_modeLiveRadio);
        modeBox.Controls.Add(_modeMockRadio);
        bottom.Controls.Add(modeBox);
        var connectBtn = new Button { Text = "Lam moi thong tin ETABS", Dock = DockStyle.Top, Height = 28 };
        connectBtn.Click += (_, _) => RefreshEtabsInfo();
        bottom.Controls.Add(connectBtn);
        page.Controls.Add(bottom);
        var aliasLabel = new Label
        {
            Dock = DockStyle.Top, AutoSize = true,
            Text = "Anh xa diem Excel -> Joint ETABS (mac dinh giong ten; sua neu joint trong model dat ten khac):",
        };
        page.Controls.Add(aliasLabel);
        var row = new Panel { Dock = DockStyle.Top, Height = 36 };
        _pointCombo.Location = new Point(0, 4);
        row.Controls.Add(_pointCombo);
        var selBtn = new Button { Text = "Lay diem dang chon trong ETABS", Location = new Point(180, 2), AutoSize = true };
        selBtn.Click += (_, _) => UseSelectedPoint();
        row.Controls.Add(selBtn);
        row.Controls.Add(new Label
        {
            Text = "(ETABS that: chon joint trong model truoc roi bam nut. Offline: dung diem mock.)",
            Location = new Point(400, 8), AutoSize = true,
        });
        page.Controls.Add(row);
        page.Controls.Add(new Label
        {
            Dock = DockStyle.Top, AutoSize = true,
            Text = "Chon diem Excel can gan (moi lan gan 1 diem; profile giu nguyen cho diem sau):",
        });
        return page;
    }

    private TabPage BuildTabAxis()
    {
        var page = new TabPage("3. Phuong & chieu");
        page.Controls.Add(_axisExampleLabel);
        var updateBtn = new Button { Text = "Cap nhat vi du", Dock = DockStyle.Top, AutoSize = true };
        updateBtn.Click += (_, _) => UpdateAxisExample();
        page.Controls.Add(updateBtn);
        var axisPanel = new Panel { Dock = DockStyle.Top, Height = 110 };
        int y = 6;
        foreach (string excelAxis in new[] { "Hx", "Hy", "Vz" })
        {
            var lbl = new Label { Text = $"{excelAxis}  ->  ETABS", Location = new Point(0, y + 4), Width = 130, AutoSize = false };
            axisPanel.Controls.Add(lbl);
            var tgt = new ComboBox { Location = new Point(140, y), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            tgt.Items.AddRange(new object[] { "X", "Y", "Z" });
            tgt.Text = excelAxis switch { "Hx" => _profile.Axis.HxAxis, "Hy" => _profile.Axis.HyAxis, _ => _profile.Axis.VzAxis };
            var sign = new ComboBox { Location = new Point(210, y), Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            sign.Items.AddRange(new object[] { "+", "-" });
            int s = excelAxis switch { "Hx" => _profile.Axis.HxSign, "Hy" => _profile.Axis.HySign, _ => _profile.Axis.VzSign };
            sign.Text = s > 0 ? "+" : "-";
            tgt.SelectedIndexChanged += (_, _) => UpdateAxisExample();
            sign.SelectedIndexChanged += (_, _) => UpdateAxisExample();
            axisPanel.Controls.Add(tgt);
            axisPanel.Controls.Add(sign);
            _axisControls[excelAxis] = (tgt, sign);
            y += 32;
        }
        page.Controls.Add(axisPanel);
        page.Controls.Add(new Label
        {
            Dock = DockStyle.Top, AutoSize = true, ForeColor = Color.FromArgb(0xA3, 0x33, 0x33),
            Text = "ETABS global +Z huong len, nen mac dinh Vz -> -Z. Dau Hx/Hy phai kiem tra 1 diem that truoc khi gan hang loat.",
        });
        page.Controls.Add(new Label
        {
            Dock = DockStyle.Top, AutoSize = true,
            Text = "Anh xa phuong/chieu — so do goc: Hx mui ten sang trai, Hy len tren, '+z down' (Vz duong khi huong xuong).",
        });
        UpdateAxisExample();
        return page;
    }

    private TabPage BuildTabMatch()
    {
        var page = new TabPage("4. Match pattern");
        page.Controls.Add(_matchPanel);
        var opts = new Panel { Dock = DockStyle.Top, Height = 36 };
        opts.Controls.Add(_createMissingCheck);
        _createMissingCheck.Location = new Point(0, 8);
        opts.Controls.Add(new Label { Text = "Hang +/-:", Location = new Point(230, 10), AutoSize = true });
        _reversibleCombo.Items.AddRange(new object[] { "plus_only", "both" });
        _reversibleCombo.Text = "both";
        _reversibleCombo.Location = new Point(300, 6);
        opts.Controls.Add(_reversibleCombo);
        opts.Controls.Add(new Label { Text = "(plus_only = chi chieu +, both = them pattern _NEG chieu am)", Location = new Point(420, 10), AutoSize = true });
        opts.Controls.Add(new Label { Text = "O khoang '-3/+6':", Location = new Point(0, 10), AutoSize = true });
        _rangeCombo.Items.AddRange(new object[] { "max_abs", "max", "min", "skip" });
        _rangeCombo.Text = "max_abs";
        _rangeCombo.Location = new Point(110, 6);
        opts.Controls.Add(_rangeCombo);
        page.Controls.Add(opts);
        var bar = new Panel { Dock = DockStyle.Top, Height = 36 };
        var suggestBtn = new Button { Text = "Goi y tu dong tat ca", Location = new Point(0, 4), AutoSize = true };
        suggestBtn.Click += (_, _) => AutoSuggestAll();
        bar.Controls.Add(suggestBtn);
        bar.Controls.Add(new Label
        {
            Text = "Goi y theo ten (BC=belt conveyor, BT=belt tension, NX/NY=notional...) — CHI LA GOI Y, ban chot lai tung dong.",
            Location = new Point(170, 8), AutoSize = true,
        });
        page.Controls.Add(bar);
        return page;
    }

    private TabPage BuildTabAssign()
    {
        var page = new TabPage("5. Preview & gan");
        page.Controls.Add(_previewBox);
        var bar = new Panel { Dock = DockStyle.Top, Height = 36 };
        var previewBtn = new Button { Text = "Xem preview", Location = new Point(0, 4), AutoSize = true };
        previewBtn.Click += (_, _) => ShowPreview();
        bar.Controls.Add(previewBtn);
        var assignBtn = new Button { Text = "GAN VAO ETABS", Location = new Point(110, 4), AutoSize = true };
        assignBtn.Click += (_, _) => Assign();
        bar.Controls.Add(assignBtn);
        var exportBtn = new Button { Text = "Xuat preview (.txt)", Location = new Point(230, 4), AutoSize = true };
        exportBtn.Click += (_, _) => ExportPreview();
        bar.Controls.Add(exportBtn);
        _replaceCheck.Location = new Point(370, 8);
        bar.Controls.Add(_replaceCheck);
        page.Controls.Add(bar);
        return page;
    }

    // ---------------- logic (port app.py) ----------------
    private void SyncProfileFromUi()
    {
        foreach (string excelAxis in new[] { "Hx", "Hy", "Vz" })
        {
            var (tgt, sign) = _axisControls[excelAxis];
            string axis = tgt.Text;
            int s = sign.Text == "+" ? +1 : -1;
            switch (excelAxis)
            {
                case "Hx": _profile.Axis.HxAxis = axis; _profile.Axis.HxSign = s; break;
                case "Hy": _profile.Axis.HyAxis = axis; _profile.Axis.HySign = s; break;
                default: _profile.Axis.VzAxis = axis; _profile.Axis.VzSign = s; break;
            }
        }
        _profile.PatternMap = _patternControls.ToDictionary(kv => kv.Key, kv => kv.Value.Text);
        _profile.PointAlias = _aliasControls.ToDictionary(kv => kv.Key, kv => kv.Value.Text.Trim());
        _profile.CreateMissingPattern = _createMissingCheck.Checked;
        _profile.ReversibleMode = _reversibleCombo.Text;
        _profile.RangePolicy = _rangeCombo.Text;
        _profile.ReplaceExisting = _replaceCheck.Checked;
    }

    private void ApplyProfileToUi()
    {
        foreach (string excelAxis in new[] { "Hx", "Hy", "Vz" })
        {
            var (tgt, sign) = _axisControls[excelAxis];
            (string axis, int s) = excelAxis switch
            {
                "Hx" => (_profile.Axis.HxAxis, _profile.Axis.HxSign),
                "Hy" => (_profile.Axis.HyAxis, _profile.Axis.HySign),
                _ => (_profile.Axis.VzAxis, _profile.Axis.VzSign),
            };
            tgt.Text = axis;
            sign.Text = s > 0 ? "+" : "-";
        }
        _createMissingCheck.Checked = _profile.CreateMissingPattern;
        _reversibleCombo.Text = _profile.ReversibleMode;
        _rangeCombo.Text = _profile.RangePolicy;
        _replaceCheck.Checked = _profile.ReplaceExisting;
        if (_table is not null)
        {
            RefreshMatchTab();
            foreach (var kv in _patternControls)
                if (_profile.PatternFor(kv.Key) != "")
                    kv.Value.Text = _profile.PatternFor(kv.Key);
            foreach (var kv in _aliasControls)
                if (_profile.PointAlias.TryGetValue(kv.Key, out string? a))
                    kv.Value.Text = a;
        }
        UpdateAxisExample();
    }

    private void TryAutoLoadProfile()
    {
        try
        {
            if (File.Exists(DefaultProfilePath))
            {
                _profile = Profile.Load(DefaultProfilePath);
                ApplyProfileToUi();
                _statusLabel.Text = $"Da tai profile lan truoc: {DefaultProfilePath} (file goc: {_profile.SourceFile}).";
            }
        }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Khong tai duoc profile lan truoc ({ex.Message}) — dung mac dinh.";
        }
    }

    private void ChooseFile()
    {
        using var dlg = new OpenFileDialog
        {
            Filter = "Excel (*.xlsx)|*.xlsx",
            Title = "Chon file Excel bang tai mong",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK)
            return;
        _fileTextBox.Text = dlg.FileName;
        _updatingSheet = true;
        try
        {
            _sheetCombo.Items.Clear();
            _sheetCombo.Items.AddRange(Detector.ListSheets(dlg.FileName).ToArray());
            if (_sheetCombo.Items.Count > 0)
                _sheetCombo.SelectedIndex = 0;
        }
        catch { _sheetCombo.Items.Clear(); }
        finally { _updatingSheet = false; }
        Detect();
    }

    private void Detect()
    {
        string path = _fileTextBox.Text.Trim();
        if (path == "" || !File.Exists(path))
        {
            MessageBox.Show(this, "Chon file Excel (.xlsx) truoc.", "Chua chon file",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        FoundationTable t;
        try
        {
            string? sheet = _sheetCombo.Text.Trim();
            if (sheet == "") sheet = null;
            t = Detector.DetectFile(path, sheet);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Khong doc duoc file",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        // Nap danh sach sheet (ca khi mo file bang cach khac, khong qua Chon file).
        if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                string cur = _sheetCombo.Text;
                _updatingSheet = true;
                _sheetCombo.Items.Clear();
                _sheetCombo.Items.AddRange(Detector.ListSheets(path).ToArray());
                if (_sheetCombo.Items.Contains(cur))
                    _sheetCombo.Text = cur;
                _updatingSheet = false;
            }
            catch { /* giu danh sach cu */ }
        }
        _table = t;
        _profile.SourceFile = path;
        _profile.Sheet = t.Sheet;
        _updatingSheet = true;
        if (t.Sheet != "" && _sheetCombo.Text == "")
            _sheetCombo.Text = t.Sheet;
        _updatingSheet = false;
        _statusLabel.Text = $"Da nhan dien: {t.Points.Count} diem, {t.Rows.Count} dong tai, {t.Patterns.Count} pattern (sheet: {t.Sheet}).";
        // Grid
        _grid.Columns.Clear();
        _grid.Rows.Clear();
        _grid.Columns.Add("Load", "Load");
        foreach (string p in t.Points)
            foreach (string comp in new[] { "Hx", "Hy", "Vz" })
                _grid.Columns.Add($"{p} {comp}", $"{p} {comp}");
        foreach (var r in t.Rows)
        {
            var vals = new List<object> { r.Name + (r.Reversible ? "  (+/-)" : "") };
            foreach (string p in t.Points)
                foreach (string comp in new[] { "Hx", "Hy", "Vz" })
                {
                    string raw = "";
                    if (r.Values.TryGetValue(p, out var compVals) && compVals.TryGetValue(comp, out var lv))
                        raw = lv.Raw;
                    vals.Add(raw);
                }
            _grid.Rows.Add(vals.ToArray());
        }
        var info = new List<string>
        {
            $"Tieu de: {(t.Title == "" ? "(khong doc duoc)" : t.Title)}",
            $"Diem: {string.Join(", ", t.Points)}",
            $"Pattern trong file ({t.Patterns.Count}): {string.Join(", ", t.Patterns.Take(40).Select(p => p.Name))}",
        };
        if (t.Warnings.Count > 0)
        {
            info.Add("Canh bao nhan dien:");
            info.AddRange(t.Warnings.Select(w => $"  ! {w}"));
        }
        _fileInfoBox.Text = string.Join("\n", info);
        // Tab diem
        _pointCombo.Items.Clear();
        _pointCombo.Items.AddRange(t.Points.ToArray());
        if (t.Points.Count > 0 && _pointCombo.Text == "")
            _pointCombo.SelectedIndex = 0;
        _aliasPanel.Controls.Clear();
        _aliasControls.Clear();
        foreach (string p in t.Points)
        {
            var row = new Panel { Height = 30, Width = 600 };
            row.Controls.Add(new Label { Text = p, Location = new Point(0, 6), Width = 90, AutoSize = false });
            var tb = new TextBox
            {
                Location = new Point(100, 3), Width = 200,
                Text = _profile.PointAlias.TryGetValue(p, out string? a) ? a : p,
            };
            row.Controls.Add(tb);
            row.Controls.Add(new Label { Text = "(joint ETABS)", Location = new Point(310, 6), AutoSize = true });
            _aliasPanel.Controls.Add(row);
            _aliasControls[p] = tb;
        }
        RefreshMatchTab();
        UpdateAxisExample();
    }

    private void RefreshMatchTab()
    {
        _matchPanel.Controls.Clear();
        _patternControls.Clear();
        if (_table is null)
            return;
        List<string> available;
        try { available = _etabs.PatternNames(); }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Khong lay duoc pattern tu ETABS: {ex.Message}";
            available = new List<string>();
        }
        var filePats = _table.Patterns.Select(p => p.Name).ToList();
        var choices = new List<string> { "" };
        choices.AddRange(new SortedSet<string>(available.Concat(filePats)));
        foreach (var r in _table.Rows)
        {
            var row = new Panel { Height = 30, Width = 700 };
            row.Controls.Add(new Label
            {
                Text = r.Name + (r.Reversible ? "  (+/-)" : ""),
                Location = new Point(0, 6), Width = 380, AutoSize = false,
            });
            var combo = new ComboBox { Location = new Point(390, 3), Width = 130, DropDownStyle = ComboBoxStyle.DropDownList };
            combo.Items.AddRange(choices.ToArray());
            string cur = _profile.PatternFor(r.Name);
            if (cur == "")
                cur = Mapping.SuggestPattern(r.Name, available);
            combo.Text = cur;
            row.Controls.Add(combo);
            _matchPanel.Controls.Add(row);
            _patternControls[r.Name] = combo;
        }
    }

    private void AutoSuggestAll()
    {
        if (_table is null)
            return;
        List<string> available;
        try { available = _etabs.PatternNames(); }
        catch { available = new List<string>(); }
        foreach (var r in _table.Rows)
            if (_patternControls.TryGetValue(r.Name, out var combo))
                combo.Text = Mapping.SuggestPattern(r.Name, available);
        _statusLabel.Text = "Da goi y tu dong — kiem tra lai tung dong truoc khi gan.";
    }

    private void UpdateAxisExample()
    {
        if (_table is not null)
            SyncProfileFromUi();
        var ax = new AxisMapping();
        foreach (string excelAxis in new[] { "Hx", "Hy", "Vz" })
        {
            var (tgt, sign) = _axisControls[excelAxis];
            int s = sign.Text == "+" ? +1 : -1;
            switch (excelAxis)
            {
                case "Hx": ax.HxAxis = tgt.Text; ax.HxSign = s; break;
                case "Hy": ax.HyAxis = tgt.Text; ax.HySign = s; break;
                default: ax.VzAxis = tgt.Text; ax.VzSign = s; break;
            }
        }
        var errs = ax.Validate();
        var (fx, fy, fz) = ax.Apply(1, -2, 19);
        _axisExampleLabel.Text =
            $"Hien tai: {ax.Describe()}\n" +
            $"Vi du — Dead loads diem 20/D (Hx=0, Hy=-2, Vz=19) -> Fx={fx:0}, Fy={fy:0}, Fz={fz:0} kN" +
            (errs.Count > 0 ? $"\nLOI: {string.Join("; ", errs)}" : "  (Vz=19 huong xuong -> Fz am la dung neu +z down).");
    }

    private void SwitchMode()
    {
        _etabs = _modeMockRadio.Checked ? _mock : _liveApi;
        _statusLabel.Text = _modeMockRadio.Checked
            ? "Che do offline (Mock ETABS)."
            : "Che do model ETABS dang mo.";
        RefreshMatchTab();
    }

    private void RefreshEtabsInfo()
    {
        try
        {
            var pats = _etabs.PatternNames();
            var pts = _etabs.ListPoints();
            _etabsInfoBox.Text = _modeMockRadio.Checked
                ? $"Mock ETABS: {pats.Count} pattern, {pts.Count} point (20/D, 20/E, 27/D, 27/E).\n"
                : $"Da ket noi model ETABS dang mo: {pats.Count} pattern, {pts.Count} point.\nPattern: {string.Join(", ", pats.Take(30))}";
            _statusLabel.Text = $"ETABS: {pats.Count} pattern, {pts.Count} point.";
            RefreshMatchTab();
        }
        catch (Exception ex)
        {
            _etabsInfoBox.Text = $"Khong lay duoc thong tin ETABS: {ex.Message}";
        }
    }

    private void UseSelectedPoint()
    {
        List<PointInfo> sel;
        try { sel = _etabs.GetSelectedPoints(); }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Loi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (sel.Count == 0)
        {
            MessageBox.Show(this,
                "Chon 1 joint trong ETABS truoc (hoac o che do Mock, diem dang chon la '20 / D').",
                "Chua chon diem", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var p = sel[0];
        _statusLabel.Text = $"Diem dang chon trong ETABS: label='{p.Label}' name='{p.Name}'. Nhap ten nay vao cot Joint ETABS o diem Excel tuong ung.";
        MessageBox.Show(this, $"Label: {p.Label}\nName: {p.Name}\nToa do: ({p.X}, {p.Y}, {p.Z})",
            "Diem dang chon", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private (string? Point, BuildResult? Result) CurrentResult()
    {
        SyncProfileFromUi();
        if (_table is null)
        {
            MessageBox.Show(this, "Import file truoc.", "Chua co bang",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return (null, null);
        }
        string point = _pointCombo.Text;
        if (point == "")
        {
            MessageBox.Show(this, "Chon diem Excel o tab 2.", "Chua chon diem",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return (null, null);
        }
        return (point, Assigner.BuildAssignments(_table, point, _profile));
    }

    private void ShowPreview()
    {
        var (point, res) = CurrentResult();
        if (res is null || point is null || _table is null)
            return;
        _previewBox.Text = Assigner.PreviewText(_table, point, _profile, res);
    }

    private void Assign()
    {
        var (point, res) = CurrentResult();
        if (res is null || point is null)
            return;
        if (res.Assignments.Count == 0)
        {
            MessageBox.Show(this, "Preview trong: chua match pattern hoac toan dong 0.",
                "Khong co gi de gan", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        string joint = _profile.PointAlias.TryGetValue(point, out string? a) ? a : point;
        var pats = res.Assignments.Select(x => x.Pattern).Distinct().OrderBy(x => x).ToList();
        if (MessageBox.Show(this,
                $"Gan {res.Assignments.Count} tai vao joint '{joint}'?\nPattern: {string.Join(", ", pats)}",
                "Xac nhan gan tai", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;
        // Tao pattern thieu (neu tick) — pattern _NEG ke thua loai cua
        // pattern goc (EX_NEG <- EX 'Seismic'), khong phai tat ca Other.
        // Port cua app.py: lay danh sach pattern hien co 1 lan roi tao tung
        // pattern thieu voi ma loai tu PatternTypes.CreateTypeFor().
        var missing = pats.Where(x => !_etabs.PatternNames().Contains(x)).ToList();
        if (missing.Count > 0)
        {
            if (_profile.CreateMissingPattern)
            {
                var existing = _etabs.ListPatterns();
                foreach (string p in missing)
                {
                    try { _etabs.CreatePattern(p, typeCode: PatternTypes.CreateTypeFor(p, existing)); }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, $"{p}: {ex.Message}", "Khong tao duoc pattern",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show(this,
                    "Pattern chua co trong ETABS: " + string.Join(", ", missing) +
                    "\nTick 'Tu tao pattern...' o tab 4, hoac tao tay trong ETABS roi gan lai.",
                    "Thieu pattern trong model", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        var errs = new List<string>();
        foreach (var assignment in res.Assignments)
        {
            try { _etabs.AssignForce(assignment.EtabsJoint, assignment.Pattern, assignment.Values(), _profile.ReplaceExisting); }
            catch (Exception ex) { errs.Add($"{assignment.Pattern}@{assignment.EtabsJoint}: {ex.Message}"); }
        }
        try { _etabs.RefreshView(); } catch { /* bo qua */ }
        if (errs.Count > 0)
            MessageBox.Show(this, string.Join("\n", errs), "Gan loi mot phan",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        else
            MessageBox.Show(this, $"Da gan {res.Assignments.Count} tai vao joint '{res.Assignments[0].EtabsJoint}'.",
                "Xong", MessageBoxButtons.OK, MessageBoxIcon.Information);
        ShowPreview();
    }

    private void ExportPreview()
    {
        var (point, res) = CurrentResult();
        if (res is null || point is null || _table is null)
            return;
        using var dlg = new SaveFileDialog
        {
            DefaultExt = ".txt",
            Filter = "Text (*.txt)|*.txt",
            FileName = $"preview_{point.Replace(" / ", "_").Replace("/", "_")}.txt",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK)
            return;
        File.WriteAllText(dlg.FileName, Assigner.PreviewText(_table, point, _profile, res));
        _statusLabel.Text = $"Da xuat preview: {dlg.FileName}";
    }

    private void SaveProfile()
    {
        SyncProfileFromUi();
        string dir = Path.GetDirectoryName(DefaultProfilePath)!;
        Directory.CreateDirectory(dir);
        using var dlg = new SaveFileDialog
        {
            DefaultExt = ".json",
            Filter = "Profile JSON (*.json)|*.json",
            InitialDirectory = dir,
            FileName = "profile.json",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK)
            return;
        _profile.Save(dlg.FileName);
        _statusLabel.Text = $"Da luu profile: {dlg.FileName} — lan sau Tai profile nay roi chi doi diem gan.";
    }

    private void LoadProfile()
    {
        string dir = Path.GetDirectoryName(DefaultProfilePath)!;
        using var dlg = new OpenFileDialog
        {
            Filter = "Profile JSON (*.json)|*.json",
            InitialDirectory = Directory.Exists(dir) ? dir : "",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK)
            return;
        try { _profile = Profile.Load(dlg.FileName); }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Khong doc duoc profile",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        ApplyProfileToUi();
        _statusLabel.Text = $"Da tai profile: {dlg.FileName} (file goc: {_profile.SourceFile}).";
    }
}
