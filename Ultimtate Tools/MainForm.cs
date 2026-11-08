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
/// MOT VIEW DUY NHAT (user chot 2026-11-08): toan bo chuc nang nam trong 1
/// panel cuon doc duy nhat — File & bang, Diem gan tai, Phuong & chieu,
/// Match pattern, Preview & gan — KHONG chia tab con, KHONG mo dialog phu.
/// (Truoc day dung TabControl 5 tab nhung header tab khong hien duoc tren
/// ETABS, ca dang embed lan dialog, nen user khong thay/bam duoc tab 2-5.)
///
/// Quy tac gan (user chot 08/11, sau khi bam GAN ma khong thay noi luc):
///   - Gan 1 diem 1 luc. Joint dich lay tu lua chon trong ETABS (ten that
///     qua nut 'Chon diem trong ETABS' / 'Lay diem dang chon') — KHONG dung
///     alias go tay (do la ly do lan truoc gan khong thay: alias '92 / D.1'
///     khong khop ten joint that trong model, SetLoadForce tra loi).
///     Truoc khi gan, tool kiem tra joint co trong model khong.
///   - Grid chi preview 1 diem khi da chon diem (khong load tat ca).
///   - Match 2 cot: trai = gia tri nho, phai = gia tri lon; chi pattern co
///     san trong model — tool KHONG tu tao pattern.
///   - Khong co che do Mock/Offline o UI — plugin chi lam viec voi file
///     dang mo. Khong co vi du dien giai o Phuong & chieu.
///
/// Lưu ý: plugin chạy trong tiến trình ETABS — KHÔNG gọi
/// ApplicationConfiguration.Initialize() / EnableVisualStyles().
/// Cài đặt lưu ra %AppData%\EtabsLoadAssigner\profile.json (cPlugin bị hủy
/// giữa các lần gọi từ menu Tools nên không giữ state trong bộ nhớ).
/// </summary>
public sealed class MainForm : Form
{
    private readonly IEtabsApi _liveApi;
    private IEtabsApi _etabs; // Luon la model ETABS dang mo — khong co che do Mock o UI (user chot 08/11).
    // Joint dich lay tu lua chon trong ETABS (ten that, khong phai alias go tay).
    private string _targetJointName = "";
    private string _targetJointLabel = "";

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
    private readonly TextBox _fileInfoBox = new() { Dock = DockStyle.Top, Height = 110, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };

    private readonly ComboBox _pointCombo = new() { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly Label _targetJointLabelCtl = new() { Dock = DockStyle.Top, AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };

    private readonly Dictionary<string, (ComboBox Axis, ComboBox Sign)> _axisControls = new();

    private readonly FlowLayoutPanel _matchPanel = new()
    {
        Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown,
        WrapContents = false, AutoScroll = true,
    };
    // Match 2 cot: moi dong tai -> (combo cot trai, combo cot phai).
    private readonly Dictionary<string, (ComboBox Left, ComboBox Right)> _patternControls = new();

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
        UpdateEtabsStatus();
    }

    // ---------------- UI ----------------
    // MOT VIEW DUY NHAT (user chot 2026-11-08): toan bo chuc nang nam trong
    // 1 panel cuon doc duy nhat, KHONG chia tab con, KHONG mo dialog phu.
    // Ly do: TabControl cua MainForm khong hien duoc header tab tren ETABS
    // (ca dang embed lan dialog) nen user khong thay/bam duoc tab 2-5.
    // Cac section xep doc tu tren xuong trong 1 Panel AutoScroll; moi
    // section la 1 GroupBox/Panel Dock=Top (them vao theo thu tu NGUOC de
    // hien dung thu tu — Dock=Top moi them se day cac section truoc xuong).
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
        _statusLabel.Text = "Chưa import file.";
        _viewTitleLabel.Text = "Chưa nhận diện bảng — chọn file Excel rồi bấm “Nhận diện bảng”.";
        var gridPanel = new Panel { Dock = DockStyle.Top, Height = 320 };
        gridPanel.Controls.Add(_grid);

        var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        // Them theo thu tu nguoc (section cuoi them truoc) de hien dung thu tu tren->duoi.
        scroll.Controls.Add(BuildSectionAssign());
        scroll.Controls.Add(BuildSectionMatch());
        scroll.Controls.Add(BuildSectionAxis());
        scroll.Controls.Add(BuildSectionPoint());
        scroll.Controls.Add(_fileInfoBox);
        scroll.Controls.Add(gridPanel);
        scroll.Controls.Add(_viewTitleLabel);
        scroll.Controls.Add(_statusLabel);
        scroll.Controls.Add(top);
        Controls.Add(scroll);
    }

    private GroupBox BuildSectionPoint()
    {
        // Gan 1 diem 1 luc (user chot 08/11). Joint dich lay tu lua chon
        // trong ETABS (ten that) — khong con danh sach alias go tay, khong
        // con che do Mock/Offline (plugin chi lam viec voi file dang mo).
        var box = new GroupBox { Text = "Điểm gán tải", Dock = DockStyle.Top, Height = 150 };
        _targetJointLabelCtl.Text = "Joint ETABS sẽ gán: (chưa chọn — bấm 'Chọn điểm trong ETABS' bên dưới).";
        box.Controls.Add(_targetJointLabelCtl);
        var row = new Panel { Dock = DockStyle.Top, Height = 36 };
        _pointCombo.Location = new Point(0, 4);
        _pointCombo.SelectedIndexChanged += (_, _) =>
        {
            // Chon diem gan -> dong bo dropdown xem (grid chi hien 1 diem nay).
            if (_pointCombo.Text != "" && _pointViewCombo.Items.Contains(_pointCombo.Text))
                _pointViewCombo.Text = _pointCombo.Text;
        };
        row.Controls.Add(_pointCombo);
        var pickBtn = new Button { Text = "Chọn điểm trong ETABS", Location = new Point(180, 2), AutoSize = true };
        pickBtn.Click += (_, _) => PickPointInEtabs();
        row.Controls.Add(pickBtn);
        var selBtn = new Button { Text = "Lấy điểm đang chọn", Location = new Point(360, 2), AutoSize = true };
        selBtn.Click += (_, _) => UseSelectedPoint();
        row.Controls.Add(selBtn);
        row.Controls.Add(new Label
        {
            Text = "('Chọn điểm trong ETABS': quay lại màn hình ETABS để chọn joint. 'Lấy điểm đang chọn': dùng joint đã chọn trước khi mở plugin.)",
            Location = new Point(540, 8), AutoSize = true,
        });
        box.Controls.Add(row);
        box.Controls.Add(new Label
        {
            Dock = DockStyle.Top, AutoSize = true,
            Text = "Chọn điểm Excel cần gán (mỗi lần gán 1 điểm; profile giữ nguyên cho điểm sau):",
        });
        return box;
    }

    private GroupBox BuildSectionAxis()
    {
        // Khong con vi du dien giai (user chot 08/11) — chi giu 3 hang anh xa + 2 dong ghi chu ngan.
        var box = new GroupBox { Text = "Phương & chiều", Dock = DockStyle.Top, Height = 200 };
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
            int sc = excelAxis switch { "Hx" => _profile.Axis.HxSign, "Hy" => _profile.Axis.HySign, _ => _profile.Axis.VzSign };
            sign.Text = sc > 0 ? "+" : "-";
            axisPanel.Controls.Add(tgt);
            axisPanel.Controls.Add(sign);
            _axisControls[excelAxis] = (tgt, sign);
            y += 32;
        }
        box.Controls.Add(axisPanel);
        box.Controls.Add(new Label
        {
            Dock = DockStyle.Top, AutoSize = true, ForeColor = Color.FromArgb(0xA3, 0x33, 0x33),
            Text = "ETABS global +Z hướng lên, nên mặc định Vz -> -Z. Dấu Hx/Hy phải kiểm tra 1 điểm thật trước khi gán hàng loạt.",
        });
        box.Controls.Add(new Label
        {
            Dock = DockStyle.Top, AutoSize = true,
            Text = "Ánh xạ phương/chiều — sơ đồ gốc: Hx mũi tên sang trái, Hy lên trên, '+z down' (Vz dương khi hướng xuống).",
        });
        return box;
    }

    private GroupBox BuildSectionMatch()
    {
        // Match 2 cot (user chot 08/11): cot trai = pattern nhan gia tri nho
        // (gia tri don / min khoang / -V hang +/-); cot phai = pattern nhan
        // gia tri lon (max khoang / +V hang +/-). Dong khong co +/- chi can
        // cot trai. Chi liet ke pattern CO SAN trong model — tool KHONG tu
        // tao pattern (bo checkbox tu tao, bo 'Hang +/-', bo 'O khoang').
        var box = new GroupBox { Text = "Match pattern", Dock = DockStyle.Top, Height = 320 };
        box.Controls.Add(_matchPanel);
        var bar = new Panel { Dock = DockStyle.Top, Height = 36 };
        var suggestBtn = new Button { Text = "Gợi ý tự động tất cả", Location = new Point(0, 4), AutoSize = true };
        suggestBtn.Click += (_, _) => AutoSuggestAll();
        bar.Controls.Add(suggestBtn);
        var refreshBtn = new Button { Text = "Làm mới pattern", Location = new Point(170, 4), AutoSize = true };
        refreshBtn.Click += (_, _) => RefreshMatchTab();
        bar.Controls.Add(refreshBtn);
        bar.Controls.Add(new Label
        {
            Text = "Gợi ý theo tên — CHỈ LÀ GỢI Ý, bạn chốt lại từng dòng. Pattern phải có sẵn trong model ETABS.",
            Location = new Point(300, 8), AutoSize = true,
        });
        box.Controls.Add(bar);
        return box;
    }

    private GroupBox BuildSectionAssign()
    {
        var box = new GroupBox { Text = "Preview & gán", Dock = DockStyle.Top, Height = 300 };
        box.Controls.Add(_previewBox);
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
        box.Controls.Add(bar);
        return box;
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
        _profile.PatternMap = _patternControls.ToDictionary(kv => kv.Key,
            kv => new RowPattern { Left = kv.Value.Left.Text, Right = kv.Value.Right.Text });
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
        _replaceCheck.Checked = _profile.ReplaceExisting;
        if (_table is not null)
        {
            RefreshMatchTab();
            foreach (var kv in _patternControls)
            {
                var rp = _profile.PatternFor(kv.Key);
                if (rp.Left != "")
                    kv.Value.Left.Text = rp.Left;
                if (rp.Right != "")
                    kv.Value.Right.Text = rp.Right;
            }
        }
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
        // Chi hien 1 diem khi da chon (user chot 08/11 — khong load tat ca
        // diem de tranh load nhieu): mac dinh chua chon, grid de trong.
        _pointViewCombo.Items.Clear();
        _pointViewCombo.Items.AddRange(t.Points.ToArray());
        _pointViewCombo.SelectedIndex = -1;
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
        // Diem gan: de trong cho user chon (khong tu chon diem dau — grid
        // chi preview khi da chon diem). Chon diem gan se dong bo dropdown xem.
        _pointCombo.Items.Clear();
        _pointCombo.Items.AddRange(t.Points.ToArray());
        _pointCombo.SelectedIndex = -1;
        RefreshMatchTab();
    }

    // Ve lai grid theo diem dang chon o dropdown "Điểm" (ngay sau dropdown
    // Sheet): CHI HIEN 1 DIEM (3 cot Hx/Hy/Vz), kem tieu de ghi ro dang xem
    // diem nao. Chua chon diem -> grid de trong (user chot 08/11 — khong
    // load tat ca diem de tranh load nhieu). Goi sau Detect() va khi doi
    // dropdown "Điểm" / diem gan.
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
        if (view == "" || !t.Points.Contains(view))
        {
            _grid.Columns.Clear();
            _grid.Rows.Clear();
            _viewTitleLabel.Text = "Chọn 1 điểm ở dropdown “Điểm” để xem bảng lực (chỉ xem 1 điểm để tránh load nhiều).";
            return;
        }
        string point = view;
        _updatingGrid = true;
        try
        {
            _grid.Columns.Clear();
            _grid.Rows.Clear();
            _grid.Columns.Add("Load", "Tải");
            _grid.Columns[0].ReadOnly = true; // Cot ten tai khong sua.
            foreach (string comp in new[] { "Hx", "Hy", "Vz" })
                _grid.Columns.Add($"{point} {comp}", $"{point} {comp}");
            foreach (var r in t.Rows)
            {
                var vals = new List<object> { r.Name + (r.Reversible ? "  (+/-)" : "") };
                foreach (string comp in new[] { "Hx", "Hy", "Vz" })
                {
                    string raw = "";
                    if (r.Values.TryGetValue(point, out var compVals) && compVals.TryGetValue(comp, out var lv))
                        raw = lv.Raw;
                    vals.Add(raw);
                }
                _grid.Rows.Add(vals.ToArray());
            }
        }
        finally { _updatingGrid = false; }
        _viewTitleLabel.Text = $"Đang xem: Điểm {point} — Bảng {t.Sheet} ({t.Rows.Count} dòng tải). Sửa trực tiếp trong bảng; giá trị mới dùng cho Preview/Gán.";
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
        // Chi pattern CO SAN trong model (sorted) + "" (khong gan) — tool
        // KHONG tu tao pattern (user chot 08/11).
        var choices = new List<string> { "" };
        choices.AddRange(new SortedSet<string>(available));
        var header = new Panel { Height = 22, Width = 710 };
        header.Controls.Add(new Label { Text = "Tải", Location = new Point(0, 3), Width = 380, AutoSize = false, Font = new Font("Segoe UI", 9F, FontStyle.Bold) });
        header.Controls.Add(new Label { Text = "Pattern — giá trị nhỏ (trái)", Location = new Point(390, 3), Width = 150, AutoSize = false, Font = new Font("Segoe UI", 9F, FontStyle.Bold) });
        header.Controls.Add(new Label { Text = "Pattern — giá trị lớn (phải)", Location = new Point(550, 3), Width = 150, AutoSize = false, Font = new Font("Segoe UI", 9F, FontStyle.Bold) });
        _matchPanel.Controls.Add(header);
        foreach (var r in _table.Rows)
        {
            var row = new Panel { Height = 30, Width = 710 };
            row.Controls.Add(new Label
            {
                Text = r.Name + (r.Reversible ? "  (+/-)" : ""),
                Location = new Point(0, 6), Width = 380, AutoSize = false,
            });
            var left = new ComboBox { Location = new Point(390, 3), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            left.Items.AddRange(choices.ToArray());
            var right = new ComboBox { Location = new Point(550, 3), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            right.Items.AddRange(choices.ToArray());
            var rp = _profile.PatternFor(r.Name);
            string curL = rp.Left;
            string curR = rp.Right;
            if (curL == "")
                curL = Mapping.SuggestPattern(r.Name, available);
            if (curR == "" && r.Reversible && curL != "")
                curR = Mapping.SuggestPatternOpposite(curL, available);
            left.Text = curL;
            right.Text = curR;
            row.Controls.Add(left);
            row.Controls.Add(right);
            _matchPanel.Controls.Add(row);
            _patternControls[r.Name] = (left, right);
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
        {
            if (!_patternControls.TryGetValue(r.Name, out var pair))
                continue;
            string left = Mapping.SuggestPattern(r.Name, available);
            pair.Left.Text = left;
            pair.Right.Text = r.Reversible && left != ""
                ? Mapping.SuggestPatternOpposite(left, available)
                : "";
        }
        _statusLabel.Text = "Đã gợi ý tự động — kiểm tra lại từng dòng trước khi gán.";
    }

    private void UpdateEtabsStatus()
    {
        try
        {
            var pats = _etabs.PatternNames();
            var pts = _etabs.ListPoints();
            _statusLabel.Text = $"ETABS: {pats.Count} pattern, {pts.Count} point.";
            RefreshMatchTab();
        }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Không lấy được thông tin ETABS: {ex.Message}";
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
                "Chưa chọn joint nào trong ETABS.\nChọn 1 joint trong model trước rồi bấm lại — hoặc bấm 'Chọn điểm trong ETABS' để quay lại màn hình ETABS chọn.",
                "Chưa chọn điểm", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        SetTargetJoint(sel[0]);
    }

    private void SetTargetJoint(PointInfo p)
    {
        _targetJointName = p.Name;
        _targetJointLabel = p.Label;
        _targetJointLabelCtl.Text = $"Joint ETABS sẽ gán: Name='{p.Name}'  Label='{p.Label}'  ({p.X:0.##}, {p.Y:0.##}, {p.Z:0.##})";
        _statusLabel.Text = $"Đã chọn joint ETABS: Name='{p.Name}' Label='{p.Label}'.";
    }

    private void PickPointInEtabs()
    {
        var parent = FindForm();
        if (parent is null)
        {
            UseSelectedPoint();
            return;
        }
        // An form chinh (modal) de user chon joint trong ETABS, roi hien lai.
        parent.Hide();
        var pick = new PickPointForm();
        pick.FormClosed += (_, _) =>
        {
            parent.Show();
            parent.Activate();
            if (pick.Confirmed)
                UseSelectedPoint();
        };
        pick.Show();
    }

    /// <summary>Cua so nho noi tren ETABS khi chon joint (modeless, TopMost) — dong thi quay lai form chinh.</summary>
    private sealed class PickPointForm : Form
    {
        public bool Confirmed { get; private set; }
        public PickPointForm()
        {
            Text = "Chọn điểm trong ETABS";
            Size = new Size(380, 150);
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = true;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            var lbl = new Label
            {
                Text = "Chọn 1 joint trong model ETABS rồi bấm 'Đã chọn'.\n(Bấm 'Hủy' để quay lại mà không chọn.)",
                Dock = DockStyle.Top, Height = 50, TextAlign = ContentAlignment.MiddleLeft,
            };
            var okBtn = new Button { Text = "Đã chọn", Location = new Point(90, 70), AutoSize = true };
            okBtn.Click += (_, _) => { Confirmed = true; Close(); };
            var cancelBtn = new Button { Text = "Hủy", Location = new Point(200, 70), AutoSize = true };
            cancelBtn.Click += (_, _) => { Confirmed = false; Close(); };
            Controls.Add(lbl);
            Controls.Add(okBtn);
            Controls.Add(cancelBtn);
        }
    }

    private (string? Point, string? Joint, BuildResult? Result) CurrentResult()
    {
        SyncProfileFromUi();
        if (_table is null)
        {
            MessageBox.Show(this, "Import file trước.", "Chưa có bảng",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return (null, null, null);
        }
        string point = _pointCombo.Text;
        if (point == "")
        {
            MessageBox.Show(this, "Chọn điểm Excel cần gán.", "Chưa chọn điểm",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return (null, null, null);
        }
        if (_targetJointName == "")
        {
            MessageBox.Show(this, "Chưa chọn joint trong ETABS.\nBấm 'Chọn điểm trong ETABS' để chọn joint cần gán (hoặc chọn joint trước khi mở plugin rồi bấm 'Lấy điểm đang chọn').",
                "Chưa chọn joint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return (null, null, null);
        }
        return (point, _targetJointName, Assigner.BuildAssignments(_table, point, _targetJointName, _profile));
    }

    private void ShowPreview()
    {
        var (point, joint, res) = CurrentResult();
        if (res is null || point is null || joint is null || _table is null)
            return;
        _previewBox.Text = Assigner.PreviewText(_table, point, joint, _profile, res);
    }

    private void Assign()
    {
        var (point, joint, res) = CurrentResult();
        if (res is null || point is null || joint is null)
            return;
        if (res.Assignments.Count == 0)
        {
            MessageBox.Show(this, "Preview trống: chưa match pattern hoặc toàn dòng 0.",
                "Không có gì để gán", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        var pats = res.Assignments.Select(x => x.Pattern).Distinct().OrderBy(x => x).ToList();
        if (MessageBox.Show(this,
                $"Gán {res.Assignments.Count} tải vào joint '{joint}' (Label='{_targetJointLabel}')?\nPattern: {string.Join(", ", pats)}",
                "Xác nhận gán tải", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;
        // Tool KHONG tu tao pattern (user chot 08/11): pattern thieu thi bao
        // loi, user tao trong ETABS roi gan lai.
        List<string> modelPats;
        try { modelPats = _etabs.PatternNames(); }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Không lấy được pattern từ ETABS: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        var missing = pats.Where(x => !modelPats.Contains(x)).ToList();
        if (missing.Count > 0)
        {
            MessageBox.Show(this,
                "Pattern chưa có trong model ETABS: " + string.Join(", ", missing) +
                "\nTạo pattern trong ETABS trước rồi gán lại (tool không tự tạo pattern).",
                "Thiếu pattern trong model", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        // Kiem tra joint co trong model (tranh gan vao joint khong ton tai — day
        // la ly do lan truoc bam GAN ma khong thay noi luc: alias go tay khong
        // khop ten joint that trong model).
        List<PointInfo> pts;
        try { pts = _etabs.ListPoints(); }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Không lấy được danh sách point từ ETABS: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (!pts.Any(p => p.Name == joint))
        {
            MessageBox.Show(this,
                $"Joint '{joint}' không có trong model ETABS đang mở.\nChọn lại joint trong ETABS rồi gán lại.",
                "Joint không tồn tại", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
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
            MessageBox.Show(this, $"Đã gán {res.Assignments.Count} tải vào joint '{joint}' (Label='{_targetJointLabel}').",
                "Xong", MessageBoxButtons.OK, MessageBoxIcon.Information);
        ShowPreview();
    }

    private void ExportPreview()
    {
        var (point, joint, res) = CurrentResult();
        if (res is null || point is null || joint is null || _table is null)
            return;
        using var dlg = new SaveFileDialog
        {
            DefaultExt = ".txt",
            Filter = "Text (*.txt)|*.txt",
            FileName = $"preview_{point.Replace(" / ", "_").Replace("/", "_")}.txt",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK)
            return;
        File.WriteAllText(dlg.FileName, Assigner.PreviewText(_table, point, joint, _profile, res));
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
