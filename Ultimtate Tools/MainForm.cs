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
/// Cửa sổ chính của plugin — port giao diện Tkinter 5 tab của app Python
/// (etabs_loader/app.py) sang WinForms. UI build bằng code (không có file
/// .Designer) để dễ review tay.
///
/// 5 tab tương ứng 5 chức năng:
///   1. File &amp; bang (tu dong nhan dien tu Excel)
///   2. Điểm gán tải (+ alias điểm Excel -> joint ETABS)
///   3. Phuong &amp; chieu Hx/Hy/Vz -> ETABS global
///   4. Match pattern (dong Excel &lt;-&gt; pattern ETABS)
///   5. Preview &amp; gan vao ETABS
///
/// Lưu ý: plugin chạy trong tiến trình ETABS — KHÔNG gọi
/// ApplicationConfiguration.Initialize() / EnableVisualStyles().
/// Cài đặt lưu ra %AppData%\EtabsLoadAssigner\profile.json (cPlugin bị hủy
/// giữa các lần gọi từ menu Tools nên không giữ state trong bộ nhớ).
/// </summary>
public sealed class MainForm : Form
{
    private readonly IEtabsApi _liveApi;
    private readonly MockEtabs _mock = new();
    private IEtabsApi _etabs;

    private FoundationTable? _table;
    private Profile _profile = new();
    private bool _updatingSheet;
    private bool _updatingGrid; // Chan de quy khi CellValueChanged chuan hoa lai o vua sua.

    private static string DefaultProfilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "EtabsLoadAssigner", "profile.json");

    // --- controls ---
    private readonly TextBox _fileTextBox = new() { Width = 260 };
    private readonly ComboBox _sheetCombo = new() { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox _pointViewCombo = new() { Width = 130, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly Label _statusLabel = new() { Dock = DockStyle.Top, Height = 26, ForeColor = Color.FromArgb(0x55, 0x55, 0x55) };
    private readonly Label _viewTitleLabel = new() { Dock = DockStyle.Top, AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };

    private readonly DataGridView _grid = new()
    {
        Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false,
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
    private readonly RadioButton _modeLiveRadio = new() { Text = "Model ETABS đang mở", Checked = true, AutoSize = true };
    private readonly RadioButton _modeMockRadio = new() { Text = "Offline — Mock ETABS (test, không cần ETABS)", AutoSize = true };

    private readonly Dictionary<string, (ComboBox Axis, ComboBox Sign)> _axisControls = new();
    private readonly Label _axisExampleLabel = new() { Dock = DockStyle.Top, AutoSize = true };

    private readonly FlowLayoutPanel _matchPanel = new()
    {
        Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown,
        WrapContents = false, AutoScroll = true,
    };
    private readonly Dictionary<string, ComboBox> _patternControls = new();
    private readonly CheckBox _createMissingCheck = new() { Text = "Tự tạo pattern chưa có trong model", AutoSize = true, Checked = true };
    private readonly ComboBox _reversibleCombo = new() { Width = 110, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox _rangeCombo = new() { Width = 110, DropDownStyle = ComboBoxStyle.DropDownList };

    private readonly CheckBox _replaceCheck = new() { Text = "Ghi đè tải cũ cùng pattern tại điểm này", Checked = true, AutoSize = true };
    private readonly TextBox _previewBox = new()
    {
        Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, WordWrap = false,
        ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10),
    };

    public MainForm(IEtabsApi liveApi)
    {
        _liveApi = liveApi;
        _etabs = liveApi;
        Text = "ETABS Load Assigner v0.1 (Plugin) — Gán tải móng từ Excel vào ETABS";
        Size = new Size(1050, 720);
        StartPosition = FormStartPosition.CenterScreen;
        BuildUi();
        _grid.CellValueChanged += Grid_CellValueChanged;
        TryAutoLoadProfile();
        RefreshEtabsInfo();
    }

    // ---------------- UI ----------------
    private void BuildUi()
    {
        var top = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
        var browseBtn = new Button { Text = "Chọn file Excel...", AutoSize = true };
        browseBtn.Click += (_, _) => ChooseFile();
        top.Controls.Add(browseBtn);
        top.Controls.Add(_fileTextBox);
        top.Controls.Add(new Label { Text = "Sheet:", AutoSize = true });
        _sheetCombo.SelectedIndexChanged += (_, _) => { if (!_updatingSheet) Detect(); };
        top.Controls.Add(_sheetCombo);
        top.Controls.Add(new Label { Text = "Điểm:", AutoSize = true });
        _pointViewCombo.SelectedIndexChanged += (_, _) => RefreshGrid();
        top.Controls.Add(_pointViewCombo);
        var detectBtn = new Button { Text = "Nhận diện bảng", AutoSize = true };
        detectBtn.Click += (_, _) => Detect();
        top.Controls.Add(detectBtn);
        var loadBtn = new Button { Text = "Tải profile...", AutoSize = true };
        loadBtn.Click += (_, _) => LoadProfile();
        top.Controls.Add(loadBtn);
        var saveBtn = new Button { Text = "Lưu profile...", AutoSize = true };
        saveBtn.Click += (_, _) => SaveProfile();
        top.Controls.Add(saveBtn);
        Controls.Add(top);
        _statusLabel.Text = "Chưa import file.";
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
        var page = new TabPage("1. File & bảng");
        page.Controls.Add(_grid);
        page.Controls.Add(_fileInfoBox);
        page.Controls.Add(_viewTitleLabel);
        _viewTitleLabel.Text = "Chưa nhận diện bảng — chọn file Excel rồi bấm “Nhận diện bảng”.";
        return page;
    }

    private TabPage BuildTabPoint()
    {
        var page = new TabPage("2. Điểm gán tải");
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
        var connectBtn = new Button { Text = "Làm mới thông tin ETABS", Dock = DockStyle.Top, Height = 28 };
        connectBtn.Click += (_, _) => RefreshEtabsInfo();
        bottom.Controls.Add(connectBtn);
        page.Controls.Add(bottom);
        var aliasLabel = new Label
        {
            Dock = DockStyle.Top, AutoSize = true,
            Text = "Ánh xạ điểm Excel -> Joint ETABS (mặc định giống tên; sửa nếu joint trong model đặt tên khác):",
        };
        page.Controls.Add(aliasLabel);
        var row = new Panel { Dock = DockStyle.Top, Height = 36 };
        _pointCombo.Location = new Point(0, 4);
        row.Controls.Add(_pointCombo);
        var selBtn = new Button { Text = "Lấy điểm đang chọn trong ETABS", Location = new Point(180, 2), AutoSize = true };
        selBtn.Click += (_, _) => UseSelectedPoint();
        row.Controls.Add(selBtn);
        row.Controls.Add(new Label
        {
            Text = "(ETABS thật: chọn joint trong model trước rồi bấm nút. Offline: dùng điểm mock.)",
            Location = new Point(400, 8), AutoSize = true,
        });
        page.Controls.Add(row);
        page.Controls.Add(new Label
        {
            Dock = DockStyle.Top, AutoSize = true,
            Text = "Chọn điểm Excel cần gán (mỗi lần gán 1 điểm; profile giữ nguyên cho điểm sau):",
        });
        return page;
    }

    private TabPage BuildTabAxis()
    {
        var page = new TabPage("3. Phương & chiều");
        page.Controls.Add(_axisExampleLabel);
        var updateBtn = new Button { Text = "Cập nhật ví dụ", Dock = DockStyle.Top, AutoSize = true };
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
            Text = "ETABS global +Z hướng lên, nên mặc định Vz -> -Z. Dấu Hx/Hy phải kiểm tra 1 điểm thật trước khi gán hàng loạt.",
        });
        page.Controls.Add(new Label
        {
            Dock = DockStyle.Top, AutoSize = true,
            Text = "Ánh xạ phương/chiều — sơ đồ gốc: Hx mũi tên sang trái, Hy lên trên, '+z down' (Vz dương khi hướng xuống).",
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
        opts.Controls.Add(new Label { Text = "Hàng +/-:", Location = new Point(230, 10), AutoSize = true });
        _reversibleCombo.Items.AddRange(new object[] { "plus_only", "both" });
        _reversibleCombo.Text = "both";
        _reversibleCombo.Location = new Point(300, 6);
        opts.Controls.Add(_reversibleCombo);
        opts.Controls.Add(new Label { Text = "(plus_only = chỉ chiều +, both = thêm pattern _NEG chiều âm)", Location = new Point(420, 10), AutoSize = true });
        opts.Controls.Add(new Label { Text = "Ô khoảng '-3/+6':", Location = new Point(0, 10), AutoSize = true });
        _rangeCombo.Items.AddRange(new object[] { "max_abs", "max", "min", "skip" });
        _rangeCombo.Text = "max_abs";
        _rangeCombo.Location = new Point(110, 6);
        opts.Controls.Add(_rangeCombo);
        page.Controls.Add(opts);
        var bar = new Panel { Dock = DockStyle.Top, Height = 36 };
        var suggestBtn = new Button { Text = "Gợi ý tự động tất cả", Location = new Point(0, 4), AutoSize = true };
        suggestBtn.Click += (_, _) => AutoSuggestAll();
        bar.Controls.Add(suggestBtn);
        bar.Controls.Add(new Label
        {
            Text = "Gợi ý theo tên (BC=belt conveyor, BT=belt tension, NX/NY=notional...) — CHỈ LÀ GỢI Ý, bạn chốt lại từng dòng.",
            Location = new Point(170, 8), AutoSize = true,
        });
        page.Controls.Add(bar);
        return page;
    }

    private TabPage BuildTabAssign()
    {
        var page = new TabPage("5. Preview & gán");
        page.Controls.Add(_previewBox);
        var bar = new Panel { Dock = DockStyle.Top, Height = 36 };
        var previewBtn = new Button { Text = "Xem preview", Location = new Point(0, 4), AutoSize = true };
        previewBtn.Click += (_, _) => ShowPreview();
        bar.Controls.Add(previewBtn);
        var assignBtn = new Button { Text = "GÁN VÀO ETABS", Location = new Point(110, 4), AutoSize = true };
        assignBtn.Click += (_, _) => Assign();
        bar.Controls.Add(assignBtn);
        var exportBtn = new Button { Text = "Xuất preview (.txt)", Location = new Point(230, 4), AutoSize = true };
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
                _statusLabel.Text = $"Đã tải profile lần trước: {DefaultProfilePath} (file gốc: {_profile.SourceFile}).";
            }
        }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Không tải được profile lần trước ({ex.Message}) — dùng mặc định.";
        }
    }

    private void ChooseFile()
    {
        using var dlg = new OpenFileDialog
        {
            Filter = "Excel (*.xlsx)|*.xlsx",
            Title = "Chọn file Excel bảng tải móng",
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
            MessageBox.Show(this, "Chọn file Excel (.xlsx) trước.", "Chưa chọn file",
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
            MessageBox.Show(this, ex.Message, "Không đọc được file",
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
        _statusLabel.Text = $"Đã nhận diện: {t.Points.Count} điểm, {t.Rows.Count} dòng tải, {t.Patterns.Count} pattern (sheet: {t.Sheet}).";
        // Dropdown xem tung diem (ngay sau dropdown Sheet o thanh tren).
        _pointViewCombo.Items.Clear();
        _pointViewCombo.Items.Add("(Tất cả điểm)");
        _pointViewCombo.Items.AddRange(t.Points.ToArray());
        _pointViewCombo.SelectedIndex = 0;
        RefreshGrid();
        var info = new List<string>
        {
            $"Tiêu đề: {(t.Title == "" ? "(không đọc được)" : t.Title)}",
            $"Điểm: {string.Join(", ", t.Points)}",
            $"Pattern trong file ({t.Patterns.Count}): {string.Join(", ", t.Patterns.Take(40).Select(p => p.Name))}",
        };
        if (t.Warnings.Count > 0)
        {
            info.Add("Cảnh báo nhận diện:");
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

    // Vẽ lại grid tab 1 theo điểm đang chọn ở dropdown "Điểm" (ngay sau
    // dropdown Sheet): "(Tất cả điểm)" -> hiện toàn bộ điểm (mỗi điểm 3 cột
    // Hx/Hy/Vz); chon 1 diem -> chi hien 3 cot cua diem do, kem tieu de ghi
    // ro dang xem diem nao. Goi sau Detect() va khi doi dropdown.
    private void RefreshGrid()
    {
        if (_table is null)
        {
            _grid.Columns.Clear();
            _grid.Rows.Clear();
            _viewTitleLabel.Text = "Chưa nhận diện bảng — chọn file Excel rồi bấm “Nhận diện bảng”.";
            return;
        }
        var t = _table;
        string view = _pointViewCombo.Text;
        bool all = view == "" || view == "(Tất cả điểm)";
        var points = all ? t.Points : new List<string> { view };
        _updatingGrid = true;
        try
        {
            _grid.Columns.Clear();
            _grid.Rows.Clear();
            _grid.Columns.Add("Load", "Tải");
            _grid.Columns[0].ReadOnly = true; // Cot ten tai khong sua.
            foreach (string p in points)
                foreach (string comp in new[] { "Hx", "Hy", "Vz" })
                    _grid.Columns.Add($"{p} {comp}", $"{p} {comp}");
            foreach (var r in t.Rows)
            {
                var vals = new List<object> { r.Name + (r.Reversible ? "  (+/-)" : "") };
                foreach (string p in points)
                    foreach (string comp in new[] { "Hx", "Hy", "Vz" })
                    {
                        string raw = "";
                        if (r.Values.TryGetValue(p, out var compVals) && compVals.TryGetValue(comp, out var lv))
                            raw = lv.Raw;
                        vals.Add(raw);
                    }
                _grid.Rows.Add(vals.ToArray());
            }
        }
        finally { _updatingGrid = false; }
        _viewTitleLabel.Text = all
            ? $"Đang xem: Tất cả {t.Points.Count} điểm — Bảng {t.Sheet} ({t.Rows.Count} dòng tải). Sửa trực tiếp trong bảng; giá trị mới dùng cho Preview/Gán."
            : $"Đang xem: Điểm {points[0]} — Bảng {t.Sheet} ({t.Rows.Count} dòng tải). Sửa trực tiếp trong bảng; giá trị mới dùng cho Preview/Gán.";
    }

    // Sua luc truc tiep trong grid: ghi nguoc vao FoundationTable (Preview &
    // Gan doc tu bang nay) roi chuan hoa hien thi theo Raw da parse.
    // Tên cột dạng "$điểm $comp" (vd "20/D Hx") — tách ở khoảng trắng cuối.
    private void Grid_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (_updatingGrid || _table is null || e.RowIndex < 0 || e.ColumnIndex <= 0)
            return;
        if (e.RowIndex >= _table.Rows.Count)
            return;
        string colName = _grid.Columns[e.ColumnIndex].Name;
        int sp = colName.LastIndexOf(' ');
        if (sp < 0)
            return;
        string point = colName.Substring(0, sp);
        string comp = colName.Substring(sp + 1);
        string newRaw = _grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
        var lv = Detector.ParseCell(newRaw);
        var row = _table.Rows[e.RowIndex];
        if (!row.Values.TryGetValue(point, out var compVals))
        {
            compVals = new Dictionary<string, LoadValue>();
            row.Values[point] = compVals;
        }
        compVals[comp] = lv;
        _updatingGrid = true;
        try { _grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = lv.Raw; }
        finally { _updatingGrid = false; }
        _statusLabel.Text = lv.Kind == LoadValueKind.Text
            ? $"Giá trị “{newRaw}” ở {row.Name} / {point} {comp} không phải số hay khoảng — sẽ bị bỏ qua khi gán."
            : $"Đã sửa {row.Name} / {point} {comp} = {lv.Raw}.";
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
            _statusLabel.Text = $"Không lấy được pattern từ ETABS: {ex.Message}";
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
        _statusLabel.Text = "Đã gợi ý tự động — kiểm tra lại từng dòng trước khi gán.";
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
            $"Hiện tại: {ax.Describe()}\n" +
            $"Ví dụ — Dead loads điểm 20/D (Hx=0, Hy=-2, Vz=19) -> Fx={fx:0}, Fy={fy:0}, Fz={fz:0} kN" +
            (errs.Count > 0 ? $"\nLỖI: {string.Join("; ", errs)}" : "  (Vz=19 hướng xuống -> Fz âm là đúng nếu +z down).");
    }

    private void SwitchMode()
    {
        _etabs = _modeMockRadio.Checked ? _mock : _liveApi;
        _statusLabel.Text = _modeMockRadio.Checked
            ? "Chế độ offline (Mock ETABS)."
            : "Chế độ model ETABS đang mở.";
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
                : $"Đã kết nối model ETABS đang mở: {pats.Count} pattern, {pts.Count} point.\nPattern: {string.Join(", ", pats.Take(30))}";
            _statusLabel.Text = $"ETABS: {pats.Count} pattern, {pts.Count} point.";
            RefreshMatchTab();
        }
        catch (Exception ex)
        {
            _etabsInfoBox.Text = $"Không lấy được thông tin ETABS: {ex.Message}";
        }
    }

    private void UseSelectedPoint()
    {
        List<PointInfo> sel;
        try { sel = _etabs.GetSelectedPoints(); }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (sel.Count == 0)
        {
            MessageBox.Show(this,
                "Chọn 1 joint trong ETABS trước (hoặc ở chế độ Mock, điểm đang chọn là '20 / D').",
                "Chưa chọn điểm", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var p = sel[0];
        _statusLabel.Text = $"Điểm đang chọn trong ETABS: label='{p.Label}' name='{p.Name}'. Nhập tên này vào cột Joint ETABS ở điểm Excel tương ứng.";
        MessageBox.Show(this, $"Label: {p.Label}\nName: {p.Name}\nTọa độ: ({p.X}, {p.Y}, {p.Z})",
            "Điểm đang chọn", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private (string? Point, BuildResult? Result) CurrentResult()
    {
        SyncProfileFromUi();
        if (_table is null)
        {
            MessageBox.Show(this, "Import file trước.", "Chưa có bảng",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return (null, null);
        }
        string point = _pointCombo.Text;
        if (point == "")
        {
            MessageBox.Show(this, "Chọn điểm Excel ở tab 2.", "Chưa chọn điểm",
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
            MessageBox.Show(this, "Preview trống: chưa match pattern hoặc toàn dòng 0.",
                "Không có gì để gán", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        string joint = _profile.PointAlias.TryGetValue(point, out string? a) ? a : point;
        var pats = res.Assignments.Select(x => x.Pattern).Distinct().OrderBy(x => x).ToList();
        if (MessageBox.Show(this,
                $"Gán {res.Assignments.Count} tải vào joint '{joint}'?\nPattern: {string.Join(", ", pats)}",
                "Xác nhận gán tải", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;
        // Tao pattern thieu (neu tick) — pattern _NEG ke thua loai cua
        // pattern goc (EX_NEG <- EX 'Seismic'), khong phai tat ca Other.
        // Port cua app.py: lấy danh sách pattern hien co 1 lan roi tao tung
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
                        MessageBox.Show(this, $"{p}: {ex.Message}", "Không tạo được pattern",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show(this,
                    "Pattern chưa có trong ETABS: " + string.Join(", ", missing) +
                    "\nTick 'Tự tạo pattern...' ở tab 4, hoặc tạo tay trong ETABS rồi gán lại.",
                    "Thiếu pattern trong model", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            MessageBox.Show(this, string.Join("\n", errs), "Gán lỗi một phần",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        else
            MessageBox.Show(this, $"Đã gán {res.Assignments.Count} tải vào joint '{res.Assignments[0].EtabsJoint}'.",
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
        _statusLabel.Text = $"Đã xuất preview: {dlg.FileName}";
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
        _statusLabel.Text = $"Đã lưu profile: {dlg.FileName} — lần sau Tải profile này rồi chỉ đổi điểm gán.";
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
            MessageBox.Show(this, ex.Message, "Không đọc được profile",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        ApplyProfileToUi();
        _statusLabel.Text = $"Đã tải profile: {dlg.FileName} (file gốc: {_profile.SourceFile}).";
    }
}
