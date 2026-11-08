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
/// panel cuon doc duy nhat — KHONG chia tab con, KHONG mo dialog phu.
/// (Truoc day dung TabControl 5 tab nhung header tab khong hien duoc tren
/// ETABS, ca dang embed lan dialog, nen user khong thay/bam duoc tab 2-5.)
///
/// LAYOUT GON (user chot 08/11, sau anh chup build 20261108-171010):
///   - Thanh tren: file / Sheet / Diem / profile. Thanh duoi CO DINH:
///     Joint ETABS + GÁN (luon thay, khong can cuon).
///   - Vung cuon giua: ten diem (tieu de bang) -> bang noi luc -> Match
///     pattern (ngay duoi bang, cung vung lam viec) -> Phuong & chieu
///     (1 hang ngang, chi set 1 lan) -> thong tin file.
///   - KHONG co Preview box / nut Xem preview / Xuat preview (.txt).
///     GÁN xong chi hien thong bao so luong (so lieu da thay trong grid).
///
/// QUY TAC GAN (user chot 08/11):
///   - Gan 1 diem 1 luc. Diem gan = diem dang chon o dropdown "Điểm" tren
///     thanh tren (cung la diem grid dang hien) — KHONG co dropdown chon
///     diem rieng o muc Diem gan tai.
///   - Joint dich chon tu DROPDOWN "Joint ETABS" (liet ke joint trong model
///     kem toa do) hoac nut "Lấy điểm đang chọn" (joint da chon truoc khi
///     mo plugin). KHONG co nut "Chọn điểm trong ETABS": form plugin la
///     MODAL (ShowDialog) nen du an form chinh, ETABS van bi khoa — user
///     khong click/chon duoc joint trong model (da user-verify 08/11).
///   - Match 2 cot: trai = gia tri nho, phai = gia tri lon; chi pattern co
///     san trong model — tool KHONG tu tao pattern.
///   - Khong co che do Mock/Offline o UI — plugin chi lam viec voi file
///     dang mo. Khong co vi du dien giai o Phuong & chieu.
///
/// BANG NOI LUC (user chot 08/11, theo anh mau "85 / E.2"):
///   - Ten diem la TIEU DE rieng tren bang (khong ghi chung vao header cot).
///   - Cot chi la: Tải | Hx [kN] | Hy [kN] | Vz [kN] (don vi [kN] o dong
///     thu 2 cua header, nhu anh mau).
///
/// PROFILE MAC DINH (user chot 08/11): tu luu + tu doc trong THU MUC CHUA
/// PLUG-IN (file LoadAssignerProfile.json canh file DLL). Mo plugin len la
/// co ngay match pattern + phuong/chieu lan truoc; doi 1 combo la tu luu.
/// Neu thu muc plug-in khong ghi duoc (vd Program Files khong quyen admin)
/// thi luu sang %AppData%\EtabsLoadAssigner\profile.json (vi tri cu) va
/// ghi chu o thanh trang thai.
///
/// Lưu ý: plugin chạy trong tiến trình ETABS — KHÔNG gọi
/// ApplicationConfiguration.Initialize() / EnableVisualStyles().
/// </summary>
public sealed class MainForm : Form
{
    private readonly IEtabsApi _liveApi;
    private IEtabsApi _etabs; // Luon la model ETABS dang mo — khong co che do Mock o UI (user chot 08/11).
    // Joint dich lay tu dropdown Joint ETABS / nut "Lấy điểm đang chọn"
    // (ten that, khong phai alias go tay).
    private string _targetJointName = "";
    private string _targetJointLabel = "";

    private FoundationTable? _table;
    private Profile _profile = new();
    private bool _updatingSheet;
    private bool _updatingGrid; // Chan de quy khi CellValueChanged chuan hoa lai o vua sua.
    private bool _updatingMatch; // Chan luu profile khi RefreshMatchTab dung lai combo.
    private bool _updatingJoint; // Chan de quy khi SetTargetJoint chon lai combo joint.
    private bool _uiReady; // Chi cho luu profile mac dinh sau khi constructor xong (tranh ghi de profile cu luc khoi tao).
    private List<PointInfo> _jointList = new();

    // --- Profile mac dinh: thu muc chua plug-in, fallback AppData ---
    private static string PluginFolder
    {
        get
        {
            try { return Path.GetDirectoryName(typeof(MainForm).Assembly.Location) ?? ""; }
            catch { return ""; }
        }
    }
    private static string PluginProfilePath =>
        PluginFolder == "" ? "" : Path.Combine(PluginFolder, "LoadAssignerProfile.json");
    private static string AppDataProfilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "EtabsLoadAssigner", "profile.json");

    // --- controls ---
    private readonly TextBox _fileTextBox = new() { Width = 260 };
    private readonly ComboBox _sheetCombo = new() { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox _pointViewCombo = new() { Width = 130, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly Label _statusLabel = new() { Dock = DockStyle.Top, Height = 26, ForeColor = Color.FromArgb(0x55, 0x55, 0x55) };
    // Ten diem lam tieu de rieng tren bang (vd "85 / E.2") — khong ghi chung vao header cot.
    private readonly Label _pointTitleLabel = new()
    {
        Dock = DockStyle.Top, Height = 28,
        Font = new Font("Segoe UI", 12F, FontStyle.Bold),
        TextAlign = ContentAlignment.MiddleCenter,
    };

    private readonly DataGridView _grid = new()
    {
        Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false,
        RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
        ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True },
        ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
        ColumnHeadersHeight = 36,
    };
    private readonly TextBox _fileInfoBox = new() { Dock = DockStyle.Top, Height = 80, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };

    private readonly ComboBox _jointCombo = new() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };

    private readonly Dictionary<string, (ComboBox Axis, ComboBox Sign)> _axisControls = new();

    private readonly FlowLayoutPanel _matchPanel = new()
    {
        Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown,
        WrapContents = false, AutoScroll = true,
    };
    // Match 2 cot: moi dong tai -> (combo cot trai, combo cot phai).
    private readonly Dictionary<string, (ComboBox Left, ComboBox Right)> _patternControls = new();

    private readonly CheckBox _replaceCheck = new() { Text = "Ghi đè tải cũ cùng pattern tại điểm này", Checked = true, AutoSize = true };

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
        _uiReady = true;
    }

    // ---------------- UI ----------------
    // MOT VIEW DUY NHAT (user chot 2026-11-08): toan bo chuc nang nam trong
    // 1 panel cuon doc duy nhat, KHONG chia tab con, KHONG mo dialog phu.
    // Thanh duoi (Joint + GÁN) co dinh ngoai vung cuon de luon thay.
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
        _pointTitleLabel.Text = "Chưa nhận diện bảng — chọn file Excel rồi bấm “Nhận diện bảng”.";
        var gridPanel = new Panel { Dock = DockStyle.Top, Height = 300 };
        gridPanel.Controls.Add(_grid);

        var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        // Them theo thu tu nguoc (section cuoi them truoc) de hien dung thu tu tren->duoi:
        // ten diem -> bang noi luc -> Match pattern -> Phuong & chieu -> thong tin file.
        scroll.Controls.Add(_fileInfoBox);
        scroll.Controls.Add(BuildSectionAxis());
        scroll.Controls.Add(BuildSectionMatch());
        scroll.Controls.Add(gridPanel);
        scroll.Controls.Add(_pointTitleLabel);

        // Thanh duoi CO DINH (ngoai vung cuon): Joint ETABS + GÁN luon thay, khong can cuon.
        var bottom = new Panel { Dock = DockStyle.Bottom, Height = 44 };
        var bottomFlow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false, AutoSize = false,
        };
        bottomFlow.Controls.Add(new Label { Text = "Joint ETABS:", AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold) });
        _jointCombo.DropDownWidth = 450; // Hien du Name/Label/toa do trong dropdown.
        _jointCombo.SelectedIndexChanged += (_, _) =>
        {
            if (_updatingJoint)
                return;
            int i = _jointCombo.SelectedIndex;
            if (i >= 0 && i < _jointList.Count)
                SetTargetJoint(_jointList[i]);
        };
        bottomFlow.Controls.Add(_jointCombo);
        var selBtn = new Button { Text = "Lấy điểm đang chọn", AutoSize = true };
        selBtn.Click += (_, _) => UseSelectedPoint();
        bottomFlow.Controls.Add(selBtn);
        _replaceCheck.CheckedChanged += (_, _) => SaveDefaultProfile();
        bottomFlow.Controls.Add(_replaceCheck);
        var assignBtn = new Button
        {
            Text = "GÁN VÀO ETABS", AutoSize = true,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
        };
        assignBtn.Click += (_, _) => Assign();
        bottomFlow.Controls.Add(assignBtn);
        bottom.Controls.Add(bottomFlow);

        // Them Fill truoc, roi Top/Bottom — Fill lay phan con lai.
        Controls.Add(scroll);
        Controls.Add(_statusLabel);
        Controls.Add(top);
        Controls.Add(bottom);
    }

    private GroupBox BuildSectionAxis()
    {
        // Phuong & chieu gon 1 hang ngang (chi set 1 lan): Hx -> [X] [-]  Hy -> [Y] [+]  Vz -> [Z] [-].
        var box = new GroupBox { Text = "Phương & chiều", Dock = DockStyle.Top, Height = 72 };
        var note = new Label
        {
            Dock = DockStyle.Top, AutoSize = true, ForeColor = Color.FromArgb(0xA3, 0x33, 0x33),
            Text = "ETABS +Z hướng lên, mặc định Vz -> -Z. Dấu Hx/Hy kiểm tra 1 điểm thật trước khi gán hàng loạt.",
        };
        var row = new Panel { Dock = DockStyle.Top, Height = 34 };
        int x = 6;
        foreach (string excelAxis in new[] { "Hx", "Hy", "Vz" })
        {
            var lbl = new Label { Text = $"{excelAxis}  ->", Location = new Point(x, 8), AutoSize = true };
            row.Controls.Add(lbl);
            x += 44;
            var tgt = new ComboBox { Location = new Point(x, 4), Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            tgt.Items.AddRange(new object[] { "X", "Y", "Z" });
            tgt.Text = excelAxis switch { "Hx" => _profile.Axis.HxAxis, "Hy" => _profile.Axis.HyAxis, _ => _profile.Axis.VzAxis };
            row.Controls.Add(tgt);
            x += 60;
            var sign = new ComboBox { Location = new Point(x, 4), Width = 50, DropDownStyle = ComboBoxStyle.DropDownList };
            sign.Items.AddRange(new object[] { "+", "-" });
            int sc = excelAxis switch { "Hx" => _profile.Axis.HxSign, "Hy" => _profile.Axis.HySign, _ => _profile.Axis.VzSign };
            sign.Text = sc > 0 ? "+" : "-";
            row.Controls.Add(sign);
            x += 55;
            _axisControls[excelAxis] = (tgt, sign);
            // Doi phuong/chieu la tu luu vao profile mac dinh (thu muc plug-in).
            tgt.SelectedIndexChanged += (_, _) => SaveDefaultProfile();
            sign.SelectedIndexChanged += (_, _) => SaveDefaultProfile();
            x += 28; // khoang cach giua 2 anh xa
        }
        // Them nguoc (note truoc, row sau) de row o tren, note o duoi.
        box.Controls.Add(note);
        box.Controls.Add(row);
        return box;
    }

    private GroupBox BuildSectionMatch()
    {
        // Match 2 cot (user chot 08/11): cot trai = pattern nhan gia tri nho
        // (gia tri don / min khoang / -V hang +/-); cot phai = pattern nhan
        // gia tri lon (max khoang / +V hang +/-). Dong khong co +/- chi can
        // cot trai. Chi liet ke pattern CO SAN trong model — tool KHONG tu
        // tao pattern (bo checkbox tu tao, bo 'Hang +/-', bo 'O khoang').
        // Dat ngay duoi bang noi luc (cung vung lam viec tren cung).
        var box = new GroupBox { Text = "Match pattern", Dock = DockStyle.Top, Height = 320 };
        box.Controls.Add(_matchPanel);
        var bar = new Panel { Dock = DockStyle.Top, Height = 36 };
        var suggestBtn = new Button { Text = "Gợi ý tự động tất cả", Location = new Point(0, 4), AutoSize = true };
        suggestBtn.Click += (_, _) => AutoSuggestAll();
        bar.Controls.Add(suggestBtn);
        var refreshBtn = new Button { Text = "Làm mới pattern", Location = new Point(170, 4), AutoSize = true };
        refreshBtn.Click += (_, _) => UpdateEtabsStatus();
        bar.Controls.Add(refreshBtn);
        bar.Controls.Add(new Label
        {
            Text = "Gợi ý theo tên — CHỈ LÀ GỢI Ý, bạn chốt lại từng dòng. Pattern phải có sẵn trong model ETABS.",
            Location = new Point(300, 8), AutoSize = true,
        });
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

    /// <summary>
    /// Profile mac dinh: tu doc khi mo plugin. Thu tu: thu muc chua plug-in
    /// (LoadAssignerProfile.json) truoc, roi AppData (vi tri cu, de tuong
    /// thich voi profile da luu truoc day).
    /// </summary>
    private void TryAutoLoadProfile()
    {
        string? found = null;
        if (PluginProfilePath != "" && File.Exists(PluginProfilePath))
            found = PluginProfilePath;
        else if (File.Exists(AppDataProfilePath))
            found = AppDataProfilePath;
        if (found is null)
            return;
        try
        {
            _profile = Profile.Load(found);
            ApplyProfileToUi();
            _statusLabel.Text = $"Đã tải profile mặc định: {found} (file gốc: {_profile.SourceFile}).";
        }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Không tải được profile mặc định ({ex.Message}) — dùng mặc định.";
        }
    }

    /// <summary>
    /// Tu luu profile mac dinh moi khi user doi match pattern / phuong-chieu
    /// / ghi de. Luu vao thu muc chua plug-in; neu khong ghi duoc thi luu
    /// sang AppData (vi tri cu) va ghi chu o thanh trang thai. Khong lam gi
    /// khi UI chua san sang (tranh ghi de profile cu luc khoi tao form).
    /// </summary>
    private void SaveDefaultProfile()
    {
        if (!_uiReady)
            return;
        try
        {
            SyncProfileFromUi();
            if (PluginProfilePath != "")
            {
                _profile.Save(PluginProfilePath);
                return;
            }
            throw new IOException("Không xác định được thư mục plug-in.");
        }
        catch (Exception)
        {
            try
            {
                string dir = Path.GetDirectoryName(AppDataProfilePath)!;
                Directory.CreateDirectory(dir);
                _profile.Save(AppDataProfilePath);
                _statusLabel.Text = "Thư mục plug-in không ghi được — đã lưu profile mặc định vào %AppData%\\EtabsLoadAssigner\\profile.json.";
            }
            catch { /* bo qua — profile chi mat khi mo lai, khong anh huong gan hien tai */ }
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
        // Diem gan = chinh diem nay (khong co dropdown chon diem rieng).
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
        RefreshMatchTab();
    }

    // Ve lai grid theo diem dang chon o dropdown "Điểm" (ngay sau dropdown
    // Sheet): CHI HIEN 1 DIEM, kem tieu de ghi ro ten diem. Chua chon diem
    // -> grid de trong (user chot 08/11 — khong load tat ca diem de tranh
    // load nhieu). Ten diem chi o TIEU DE tren bang, cot chi la
    // Tải | Hx [kN] | Hy [kN] | Vz [kN] (user chot 08/11 theo anh mau).
    private void RefreshGrid()
    {
        if (_table is null)
        {
            _grid.Columns.Clear();
            _grid.Rows.Clear();
            _pointTitleLabel.Text = "Chưa nhận diện bảng — chọn file Excel rồi bấm “Nhận diện bảng”.";
            return;
        }
        var t = _table;
        string view = _pointViewCombo.Text;
        if (view == "" || !t.Points.Contains(view))
        {
            _grid.Columns.Clear();
            _grid.Rows.Clear();
            _pointTitleLabel.Text = "Chọn 1 điểm ở dropdown “Điểm” để xem bảng lực (chỉ xem 1 điểm để tránh load nhiều).";
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
            // Header 2 dong: Hx / [kN] (nhu anh mau "85 / E.2").
            _grid.Columns.Add("Hx", "Hx\n[kN]");
            _grid.Columns.Add("Hy", "Hy\n[kN]");
            _grid.Columns.Add("Vz", "Vz\n[kN]");
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
        _pointTitleLabel.Text = point;
        _statusLabel.Text = $"Đang xem điểm {point} — bảng {t.Sheet} ({t.Rows.Count} dòng tải). Sửa trực tiếp trong bảng; giá trị mới dùng cho Gán.";
    }

    // Sua luc truc tiep trong grid: ghi nguoc vao FoundationTable (Gan doc
    // tu bang nay) roi chuan hoa hien thi theo Raw da parse.
    // Ten cot la ten thanh phan ("Hx"/"Hy"/"Vz"); diem lay tu dropdown
    // "Điểm" tren thanh tren (grid chi hien 1 diem).
    private void Grid_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (_updatingGrid || _table is null || e.RowIndex < 0 || e.ColumnIndex <= 0)
            return;
        if (e.RowIndex >= _table.Rows.Count)
            return;
        string comp = _grid.Columns[e.ColumnIndex].Name;
        string point = _pointViewCombo.Text;
        if (point == "")
            return;
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
        _updatingMatch = true;
        try
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
                // Doi match la tu luu vao profile mac dinh (thu muc plug-in).
                left.SelectedIndexChanged += (_, _) => { if (!_updatingMatch) SaveDefaultProfile(); };
                right.SelectedIndexChanged += (_, _) => { if (!_updatingMatch) SaveDefaultProfile(); };
                row.Controls.Add(left);
                row.Controls.Add(right);
                _matchPanel.Controls.Add(row);
                _patternControls[r.Name] = (left, right);
            }
        }
        finally { _updatingMatch = false; }
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
        SaveDefaultProfile();
    }

    private void UpdateEtabsStatus()
    {
        try
        {
            var pats = _etabs.PatternNames();
            var pts = _etabs.ListPoints();
            _statusLabel.Text = $"ETABS: {pats.Count} pattern, {pts.Count} point.";
            RefreshJointCombo();
            RefreshMatchTab();
        }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Không lấy được thông tin ETABS: {ex.Message}";
        }
    }

    /// <summary>
    /// Nap dropdown Joint ETABS: tat ca joint trong model, xep theo Z (thap
    /// truoc — joint mong o duoi), roi X, Y. Hien Name/Label/toa do de user
    /// nhan ra joint can gan (vd joint o giao 92/D.1). Giu lai joint dang
    /// chon neu van con trong model.
    /// </summary>
    private void RefreshJointCombo()
    {
        List<PointInfo> pts;
        try { pts = _etabs.ListPoints(); }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Không lấy được danh sách point từ ETABS: {ex.Message}";
            return;
        }
        _jointList = pts.OrderBy(p => p.Z).ThenBy(p => p.X).ThenBy(p => p.Y).ThenBy(p => p.Name).ToList();
        string cur = _targetJointName;
        _updatingJoint = true;
        try
        {
            _jointCombo.Items.Clear();
            foreach (var p in _jointList)
                _jointCombo.Items.Add($"Name='{p.Name}'  Label='{p.Label}'  ({p.X:0.##}, {p.Y:0.##}, {p.Z:0.##})");
            int idx = cur == "" ? -1 : _jointList.FindIndex(p => p.Name == cur);
            _jointCombo.SelectedIndex = idx;
        }
        finally { _updatingJoint = false; }
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
                "Chưa chọn joint nào trong ETABS.\nChọn 1 joint trong model ETABS trước khi mở plugin rồi bấm lại — hoặc chọn joint ở dropdown 'Joint ETABS' bên dưới.",
                "Chưa chọn điểm", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        SetTargetJoint(sel[0]);
    }

    private void SetTargetJoint(PointInfo p)
    {
        _targetJointName = p.Name;
        _targetJointLabel = p.Label;
        _statusLabel.Text = $"Đã chọn joint ETABS: Name='{p.Name}' Label='{p.Label}' ({p.X:0.##}, {p.Y:0.##}, {p.Z:0.##}).";
        // Dong bo dropdown Joint ETABS (neu joint co trong danh sach).
        _updatingJoint = true;
        try
        {
            int idx = _jointList.FindIndex(x => x.Name == p.Name);
            if (idx >= 0)
                _jointCombo.SelectedIndex = idx;
            else
                _jointCombo.SelectedIndex = -1;
        }
        finally { _updatingJoint = false; }
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
        // Diem gan = diem dang chon o dropdown "Điểm" (cung la diem grid dang hien).
        string point = _pointViewCombo.Text;
        if (point == "" || !_table.Points.Contains(point))
        {
            MessageBox.Show(this, "Chọn điểm Excel cần gán ở dropdown “Điểm” trên thanh trên.", "Chưa chọn điểm",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return (null, null, null);
        }
        if (_targetJointName == "")
        {
            MessageBox.Show(this, "Chưa chọn joint trong ETABS.\nChọn joint ở dropdown 'Joint ETABS' bên dưới (hoặc chọn joint trong model trước khi mở plugin rồi bấm 'Lấy điểm đang chọn').",
                "Chưa chọn joint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return (null, null, null);
        }
        return (point, _targetJointName, Assigner.BuildAssignments(_table, point, _targetJointName, _profile));
    }

    private void Assign()
    {
        var (point, joint, res) = CurrentResult();
        if (res is null || point is null || joint is null)
            return;
        if (res.Assignments.Count == 0)
        {
            MessageBox.Show(this, "Không có gì để gán: chưa match pattern hoặc toàn dòng 0.",
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
    }

    private void SaveProfile()
    {
        SyncProfileFromUi();
        string dir = PluginFolder != "" ? PluginFolder : Path.GetDirectoryName(AppDataProfilePath)!;
        using var dlg = new SaveFileDialog
        {
            DefaultExt = ".json",
            Filter = "Profile JSON (*.json)|*.json",
            InitialDirectory = Directory.Exists(dir) ? dir : "",
            FileName = "LoadAssignerProfile.json",
        };
        if (dlg.ShowDialog(this) != DialogResult.OK)
            return;
        _profile.Save(dlg.FileName);
        _statusLabel.Text = $"Đã lưu profile: {dlg.FileName}.";
    }

    private void LoadProfile()
    {
        string dir = PluginFolder != "" ? PluginFolder : Path.GetDirectoryName(AppDataProfilePath)!;
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
