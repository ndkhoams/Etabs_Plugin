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
/// LAYOUT GON (user chot 08/11, sau anh chup build 20261108-184715):
///   - Thanh tren: file / File gan nhat / Sheet / Diem — dung
///     TableLayoutPanel, moi control can giua doc trong hang (thang hang,
///     can doi). Nut "File gần nhất" mo lai file Excel mo lan cuoi (luu
///     duong dan vao file text nho canh DLL). KHONG co nut "Nhận diện
///     bảng" (thua — mo file Excel la tu nhan dien).
///   - Thanh duoi: KHONG CON (user chot 08/11 — cum nut Ghi de + "Lưu data
///     và Excel" + "GÁN VÀO ETABS" chuyen len hang tieu de de tiet kiem
///     dien tich). Dropdown Joint ETABS don cung da xoa (joint chon truc
///     tiep trong bang load point).
///   - Vung cuon giua: hang tieu de (trai: 2 nut chuyen bang + "Thêm điểm"
///     + "Xoá dòng" + "Làm trắng bảng"; phai: Ghi de + "Lưu data và Excel"
///     + "GÁN VÀO ETABS"; giua: ten diem) -> bang noi luc / bang load point
///     (gom ca 2 cot Match pattern ngay sau cot Tai: Pattern gia tri 1 /
///     gia tri 2) -> Phuong & chieu (cao 80 — tang them de chu khoi bi sat).
///   - GAN HANG LOAT (user chot 08/11): nut GÁN gan TAT CA cap (Diem Excel
///     -> Joint ETABS) trong bang load point, moi diem dung chung profile
///     match pattern + phuong/chieu. Chuyen bang bang 2 nut o hang tieu de
///     ("Bảng load pattern" / "Bảng load point"); them diem vao bang bang
///     nut "Thêm điểm" (lay Diem Excel dang chon o dropdown tren, cot Joint
///     de trong de chon trong bang) hoac go truc tiep trong bang; nut "Xoá
///     dòng" xoa dong dang chon (de xoa dong thua); nut "Làm trắng bảng"
///     xoa het dong de gan luot moi.
///   - KHONG co box thong tin file (user chot 08/11 — xoa).
///     KHONG co Preview box / nut Xem preview / Xuat preview (.txt).
///     GÁN xong chi hien thong bao so luong (so lieu da thay trong grid).
///
/// NHAN DIEN BANG (user chot 08/11): mo file Excel la tu nhan dien (khong
/// can bam nut). Sheet dang chon khong co data (vd README — sheet dau tien)
/// thi tu dong chuyen sang sheet tiep theo co data; chi khi CA FILE khong
/// co data moi hien hop thoai bao loi.
///
/// QUY TAC GAN (user chot 08/11):
///   - Gan 1 diem 1 luc. Diem gan = diem dang chon o dropdown "Điểm" tren
///     thanh tren (cung la diem grid dang hien) — KHONG co dropdown chon
///     diem rieng o muc Diem gan tai.
///   - Joint dich chon trong DROPDOWN "Joint ETABS" o thanh duoi (chi liet
///     ke joint DANG CHON trong ETABS — khong phai toan bo joint cua model,
///     user chot 08/11). Mo plugin len la dropdown nap san joint dang chon.
///     KHONG co nut "Chọn điểm trong ETABS" / "Lấy điểm đang chọn": form
///     plugin la MODAL (ShowDialog) nen du an form chinh, ETABS van bi khoa
///     — user khong click/chon duoc joint trong model (da user-verify
///     08/11); muon doi joint dich thi dong plugin, chon lai joint trong
///     model, mo lai plugin.
///   - Match 2 cot: trai = gia tri 1, phai = gia tri 2 (chi khi o co 2
///     gia tri); chi pattern co san trong model — tool KHONG tu tao pattern.
///   - Ghi de (checkbox "Xoá trắng tải cũ tại điểm này trước khi gán",
///     user chot 08/11): truoc khi gan moi, XOA TRANG moi tai cu tai joint
///     — ca tai khong co trong file Excel cung bi xoa, de diem chi con tai
///     do plugin nhap vao. Khong tick thi chi ghi de/ghi them tung pattern.
///   - Khong co che do Mock/Offline o UI — plugin chi lam viec voi file
///     dang mo. Khong co vi du dien giai o Phuong & chieu.
///
/// BANG NOI LUC (user chot 08/11, theo anh mau "85 / E.2"):
///   - Ten diem la TIEU DE rieng tren bang (khong ghi chung vao header cot).
///   - Cot: Tải | Pattern — giá trị 1 (trái) | Pattern — giá trị 2 (phải)
///     | Hx [kN] | Hy [kN] | Vz [kN] (don vi [kN] o dong thu 2 cua header,
///     nhu anh mau). 2 cot Pattern nam ngay sau cot Tai (user chot 08/11,
///     anh build 20261108-184715) — moi dong tai hien luon pattern cua no,
///     khong can panel Match rieng ben duoi bang.
///   - Bang KEO DAI cho kin vung cuon (user chot 08/11, anh build
///     20261108-191050) — khong de khoang trong duoi muc Phuong &amp; chieu
///     bi lang phi. Cac cot Fill het chieu rong bang: Tai 28, Pattern 1 =
///     26, Pattern 2 = 26 (bang nhau, du thay het ten pattern), Hx/Hy/Vz
///     7 moi cot (rong hon truoc de thay ro gia tri).
///
/// LUU MATCH (user chot 08/11): KHONG dung file .json rieng nua — thong tin
/// match (pattern 2 cot / phuong-chieu / ghi de) luu luon vao SHEET
/// "LoadAssigner" trong chinh file Excel noi luc. Mo file len la doc sheet
/// do (co thi nap, khong co thi dung mac dinh); bam nut "Lưu match vào
/// Excel" moi ghi lai sheet (luu THU CONG, khong tu dong — tu dong luu
/// truoc day de ghi sheet TRONG khi bang chua nhan dien xong, vd sheet
/// README 0 dong). File dang mo trong Excel thi khong ghi duoc — bao o
/// thanh trang thai, dong file roi bam Luu lai.
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
    private FoundationTable? _table;
    private Profile _profile = new();
    private bool _updatingSheet;
    private bool _updatingGrid; // Chan de quy khi CellValueChanged chuan hoa lai o vua sua.
    private bool _uiReady; // Chi cho luu profile mac dinh sau khi constructor xong (tranh ghi de profile cu luc khoi tao).
    private List<PointInfo> _jointList = new(); // Joint dang chon trong ETABS (nap vao cot Joint o bang load point).

    // --- File Excel mo lan cuoi (nut chon nhanh, user chot 08/11) ---
    // Luu duong dan file Excel mo lan cuoi vao file text nho canh DLL
    // (fallback %AppData% neu thu muc plug-in khong ghi duoc). Chi luu
    // DUONG DAN file (khong phai thong tin match — match luu trong sheet
    // "LoadAssigner" cua file Excel).
    private static string PluginFolder
    {
        get
        {
            try { return Path.GetDirectoryName(typeof(MainForm).Assembly.Location) ?? ""; }
            catch { return ""; }
        }
    }
    private static string RecentFilePath =>
        PluginFolder == "" ? "" : Path.Combine(PluginFolder, "LoadAssignerRecent.txt");
    private static string AppDataRecentPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "EtabsLoadAssigner", "recent.txt");

    private static string? LoadRecentFile()
    {
        foreach (string p in new[] { RecentFilePath, AppDataRecentPath })
        {
            try
            {
                if (p != "" && File.Exists(p))
                {
                    string s = File.ReadAllText(p).Trim();
                    if (s != "" && File.Exists(s))
                        return s;
                }
            }
            catch { /* bo qua — khong co file gan nhat */ }
        }
        return null;
    }

    private static void SaveRecentFile(string path)
    {
        try
        {
            if (RecentFilePath != "")
            {
                File.WriteAllText(RecentFilePath, path);
                return;
            }
            throw new IOException("Không xác định được thư mục plug-in.");
        }
        catch
        {
            try
            {
                string dir = Path.GetDirectoryName(AppDataRecentPath)!;
                Directory.CreateDirectory(dir);
                File.WriteAllText(AppDataRecentPath, path);
            }
            catch { /* bo qua — chi mat nut chon nhanh lan sau */ }
        }
    }

    // --- controls ---
    private readonly TextBox _fileTextBox = new() { Width = 260 };
    private readonly ComboBox _sheetCombo = new() { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox _pointViewCombo = new() { Width = 130, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly Label _statusLabel = new() { Dock = DockStyle.Top, Height = 26, ForeColor = Color.FromArgb(0x55, 0x55, 0x55) };
    // Ten diem lam tieu de rieng tren bang (vd "85 / E.2") — khong ghi chung vao header cot.
    // Nam trong _titleBar (hang tieu de, cao 30) cung 2 nut chuyen bang +
    // nut "Thêm điểm"; label gian het phan con lai, chu can giua.
    private readonly Label _pointTitleLabel = new()
    {
        Dock = DockStyle.Fill,
        Font = new Font("Segoe UI", 12F, FontStyle.Bold),
        TextAlign = ContentAlignment.MiddleCenter,
    };

    private readonly DataGridView _grid = new()
    {
        Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false,
        RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
        ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True },
        ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
        ColumnHeadersHeight = 36,
    };

    // Bang phu LOAD POINT (user chot 08/11 — gan nhieu diem 1 luc): moi dong
    // la 1 cap (Diem Excel -> Joint ETABS). User them dong truc tiep trong
    // bang (chon diem/joint o dropdown trong o) hoac bam nut "Thêm điểm".
    // Nut GÁN gan TAT CA dong trong bang nay (khong con gan 1 diem nua).
    private readonly DataGridView _pointGrid = new()
    {
        Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = true,
        AllowUserToDeleteRows = true,
        RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
        ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True },
        ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
        ColumnHeadersHeight = 36,
        Visible = false,
    };

    private readonly Dictionary<string, (ComboBox Axis, ComboBox Sign)> _axisControls = new();

    private readonly CheckBox _replaceCheck = new() { Text = "Xoá trắng tải cũ tại điểm này trước khi gán", Checked = true, AutoSize = true };

    // Vung cuon + cac muc ben trong (de keo dai bang noi luc cho kin vung
    // cuon khi cua so rong — user chot 08/11, anh build 20261108-191050).
    private Panel? _scrollPanel;
    private Panel? _gridPanel;
    private GroupBox? _axisSection;
    private Panel? _titleBar; // Hang tieu de: 2 nut chuyen bang + ten diem o giua.

    public MainForm(IEtabsApi liveApi)
    {
        _liveApi = liveApi;
        _etabs = liveApi;
        Text = "ETABS Load Assigner v0.1 (Plugin) — Gán tải móng từ Excel vào ETABS";
        Size = new Size(1050, 720);
        StartPosition = FormStartPosition.CenterScreen;
        BuildUi();
        _grid.CellValueChanged += Grid_CellValueChanged;
        // Gia tri pattern trong profile co the khong con trong model ETABS
        // (model doi) — hien o do thay vi nem loi khi set gia tri cell.
        _grid.DataError += (_, e) => { e.ThrowException = false; };
        // Mo form len: keo dai bang noi luc cho kin vung cuon.
        Load += (_, _) => LayoutGridHeight();
        UpdateEtabsStatus();
        _uiReady = true;
    }

    // ---------------- UI ----------------
    // MOT VIEW DUY NHAT (user chot 2026-11-08): toan bo chuc nang nam trong
    // 1 panel cuon doc duy nhat, KHONG chia tab con, KHONG mo dialog phu.
    // Cum nut GÁN (Ghi de + Luu + GÁN) nam o hang tieu de, luon thay.
    private void BuildUi()
    {
        // Thanh tren: TableLayoutPanel 1 hang — moi control can giua doc
        // trong o cua no (thang hang, can doi). Cot duong dan file gian
        // theo chieu rong con lai; cac cot khac vua noi dung.
        var top = new TableLayoutPanel
        {
            Dock = DockStyle.Top, AutoSize = true, ColumnCount = 7, RowCount = 1,
            Padding = new Padding(3, 3, 3, 3),
        };
        top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Chon file
        top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // File gan nhat
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); // Duong dan (gian)
        top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Nhan Sheet
        top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Combo Sheet
        top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Nhan Diem
        top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Combo Diem
        top.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        var browseBtn = new Button { Text = "Chọn file Excel...", AutoSize = true, Anchor = AnchorStyles.None };
        browseBtn.Click += (_, _) => ChooseFile();
        top.Controls.Add(browseBtn, 0, 0);
        // Nut chon nhanh file Excel mo lan cuoi (user chot 08/11): bam la mo
        // lai file do ngay, khoi browse lai. Mo khi khong co file gan nhat
        // (hoac file da bi xoa/doi ten). Tooltip hien duong dan day du.
        var recentBtn = new Button { Text = "File gần nhất", AutoSize = true, Anchor = AnchorStyles.None };
        string? recent = LoadRecentFile();
        recentBtn.Enabled = recent is not null;
        if (recent is not null)
        {
            recentBtn.Click += (_, _) => OpenRecentFile();
            var tip = new ToolTip();
            tip.SetToolTip(recentBtn, recent);
        }
        top.Controls.Add(recentBtn, 1, 0);
        _fileTextBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        top.Controls.Add(_fileTextBox, 2, 0);
        top.Controls.Add(new Label { Text = "Sheet:", AutoSize = true, Anchor = AnchorStyles.None }, 3, 0);
        _sheetCombo.Anchor = AnchorStyles.None;
        _sheetCombo.SelectedIndexChanged += (_, _) => { if (!_updatingSheet) Detect(); };
        top.Controls.Add(_sheetCombo, 4, 0);
        top.Controls.Add(new Label { Text = "Điểm:", AutoSize = true, Anchor = AnchorStyles.None }, 5, 0);
        _pointViewCombo.Anchor = AnchorStyles.None;
        _pointViewCombo.SelectedIndexChanged += (_, _) => RefreshGrid();
        top.Controls.Add(_pointViewCombo, 6, 0);
        _statusLabel.Text = "Chưa import file.";
        _pointTitleLabel.Text = "Chưa nhận diện bảng — chọn file Excel để tự nhận diện bảng.";
        _gridPanel = new Panel { Dock = DockStyle.Top, Height = 300 };
        _gridPanel.Controls.Add(_grid);
        _gridPanel.Controls.Add(_pointGrid);
        _axisSection = BuildSectionAxis();

        // Hang tieu de (user chot 08/11): trai = 2 nut chuyen bang + "Thêm
        // điểm" + "Làm trắng bảng"; phai = Ghi de + "Lưu data và Excel" +
        // "GÁN VÀO ETABS" (cum nut duoi cu chuyen len day de tiet kiem dien
        // tich — bo han thanh duoi, user chot 08/11); giua = ten diem.
        // (Dock Left: them sau thi nam trai hon; Dock Right: them sau thi
        // nam phai hon; label Fill them truoc de gian phan con lai.)
        _titleBar = new Panel { Dock = DockStyle.Top, Height = 34 };
        _titleBar.Controls.Add(_pointTitleLabel);
        var clearBtn = new Button { Text = "Làm trắng bảng", Dock = DockStyle.Left, AutoSize = true };
        clearBtn.Click += (_, _) => ClearPointGrid();
        _titleBar.Controls.Add(clearBtn);
        var delRowBtn = new Button { Text = "Xoá dòng", Dock = DockStyle.Left, AutoSize = true };
        delRowBtn.Click += (_, _) => DeleteSelectedPointRows();
        _titleBar.Controls.Add(delRowBtn);
        var addPointBtn = new Button { Text = "Thêm điểm", Dock = DockStyle.Left, AutoSize = true };
        addPointBtn.Click += (_, _) => AddCurrentPointToPointGrid();
        _titleBar.Controls.Add(addPointBtn);
        var pointBtn = new Button { Text = "Bảng load point", Dock = DockStyle.Left, AutoSize = true };
        pointBtn.Click += (_, _) => ShowPointGrid();
        _titleBar.Controls.Add(pointBtn);
        var patternBtn = new Button { Text = "Bảng load pattern", Dock = DockStyle.Left, AutoSize = true };
        patternBtn.Click += (_, _) => ShowPatternGrid();
        _titleBar.Controls.Add(patternBtn);
        _replaceCheck.Dock = DockStyle.Right;
        _titleBar.Controls.Add(_replaceCheck);
        var saveBtn = new Button { Text = "Lưu data và Excel", Dock = DockStyle.Right, AutoSize = true };
        saveBtn.Click += (_, _) => SaveMatchToExcel();
        _titleBar.Controls.Add(saveBtn);
        var assignBtn = new Button
        {
            Text = "GÁN VÀO ETABS", Dock = DockStyle.Right, AutoSize = true,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
        };
        assignBtn.Click += (_, _) => Assign();
        _titleBar.Controls.Add(assignBtn);

        _scrollPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        // Them theo thu tu nguoc (section cuoi them truoc) de hien dung thu tu tren->duoi:
        // hang tieu de (nut chuyen bang + ten diem + cum nut GÁN) -> bang noi luc / bang load point -> Phuong & chieu.
        // (Thanh duoi da bo — cum nut Ghi de/Lưu/GÁN chuyen len hang tieu de, user chot 08/11.)
        _scrollPanel.Controls.Add(_axisSection);
        _scrollPanel.Controls.Add(_gridPanel);
        _scrollPanel.Controls.Add(_titleBar);
        // Cua so doi kich thuoc -> keo dai bang noi luc cho kin vung cuon.
        _scrollPanel.Resize += (_, _) => LayoutGridHeight();

        // Them Fill truoc, roi Top — Fill lay phan con lai. (Khong con thanh
        // duoi — cum nut GÁN da len hang tieu de.)
        Controls.Add(_scrollPanel);
        Controls.Add(_statusLabel);
        Controls.Add(top);
    }

    /// <summary>
    /// Keo dai bang noi luc cho kin vung cuon (user chot 08/11 — khong de
    /// khoang trong duoi muc Phuong &amp; chieu bi lang phi). Chieu cao bang
    /// = chieu cao vung cuon tru cac muc khac (tieu de, Phuong &amp; chieu);
    /// toi thieu 200 de cua so nho van thay duoc bang.
    /// </summary>
    private void LayoutGridHeight()
    {
        if (_scrollPanel is null || _gridPanel is null || _axisSection is null || _titleBar is null)
            return;
        int other = _titleBar.Height + _axisSection.Height;
        int avail = _scrollPanel.ClientSize.Height - other;
        _gridPanel.Height = Math.Max(200, avail);
    }

    // ---------------- bang load point (gan nhieu diem 1 luc) ----------------
    // 2 nut chuyen bang o hang tieu de (user chot 08/11): "Bảng load pattern"
    // hien bang match pattern (mac dinh); "Bảng load point" hien bang phu
    // (Diem Excel -> Joint ETABS). Chi 1 bang hien tai 1 thoi diem.
    private void ShowPatternGrid()
    {
        _pointGrid.Visible = false;
        _grid.Visible = true;
        _pointTitleLabel.Text = _pointViewCombo.Text == ""
            ? "Chọn 1 điểm ở dropdown “Điểm” để xem bảng lực (chỉ xem 1 điểm để tránh load nhiều)."
            : _pointViewCombo.Text;
    }

    private void ShowPointGrid()
    {
        _grid.Visible = false;
        _pointGrid.Visible = true;
        RefreshPointGrid();
        int n = PointGridMappings().Count;
        _pointTitleLabel.Text = n == 0
            ? "Bảng load point — chưa có điểm nào."
            : $"Bảng load point — {n} điểm.";
    }

    /// <summary>
    /// Lap/dong bo 2 cot dropdown cua bang load point: cot Diem Excel lay tu
    /// _table.Points, cot Joint ETABS lay tu joint DANG CHON trong ETABS
    /// (_jointList). Giu nguyen cac dong user da nhap (chi cap nhat Items).
    /// </summary>
    private void RefreshPointGrid()
    {
        if (_pointGrid.Columns.Count == 0)
        {
            var ptCol = new DataGridViewComboBoxColumn
            {
                Name = "ExcelPoint", HeaderText = "Điểm Excel",
                DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox,
                FillWeight = 50, MinimumWidth = 120,
            };
            var jtCol = new DataGridViewComboBoxColumn
            {
                Name = "EtabsJoint", HeaderText = "Joint ETABS",
                DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox,
                FillWeight = 50, MinimumWidth = 200,
            };
            _pointGrid.Columns.Add(ptCol);
            _pointGrid.Columns.Add(jtCol);
            _pointGrid.DataError += (_, e) => { e.ThrowException = false; };
        }
        var pt = (DataGridViewComboBoxColumn)_pointGrid.Columns["ExcelPoint"];
        var jt = (DataGridViewComboBoxColumn)_pointGrid.Columns["EtabsJoint"];
        pt.Items.Clear();
        if (_table is not null)
            foreach (string p in _table.Points)
                pt.Items.Add(p);
        jt.Items.Clear();
        foreach (var p in _jointList)
            jt.Items.Add(JointDisplay(p));
    }

    private static string JointDisplay(PointInfo p) =>
        $"Name='{p.Name}'  Label='{p.Label}'  ({p.X:0.##}, {p.Y:0.##}, {p.Z:0.##})";

    /// <summary>
    /// Doc cac cap (Diem Excel, ten Joint) da nhap trong bang load point.
    /// Bo qua dong moi (placeholder), dong thieu diem hoac thieu joint.
    /// </summary>
    private List<(string ExcelPoint, string JointName)> PointGridMappings()
    {
        var list = new List<(string, string)>();
        foreach (DataGridViewRow row in _pointGrid.Rows)
        {
            if (row.IsNewRow)
                continue;
            string pt = row.Cells["ExcelPoint"].Value?.ToString() ?? "";
            string jt = row.Cells["EtabsJoint"].Value?.ToString() ?? "";
            if (pt == "" || jt == "")
                continue;
            // Cot Joint luu dang hien thi "Name='13'  Label='3'  (20.67, 0, 1.3)" — lay ten joint that.
            string jointName = ParseJointName(jt);
            if (jointName == "")
                continue;
            list.Add((pt, jointName));
        }
        return list;
    }

    private static string ParseJointName(string display)
    {
        // "Name='13'  Label='3'  (...)" -> "13". Neu user go ten truc tiep thi lay nguyen chuoi.
        int s = display.IndexOf("Name='", StringComparison.Ordinal);
        if (s >= 0)
        {
            int e = display.IndexOf('\'', s + 6);
            if (e > s + 6)
                return display.Substring(s + 6, e - (s + 6));
        }
        return display.Trim();
    }

    /// <summary>
    /// Them 1 dong moi vao bang load point: lay Diem Excel dang chon o
    /// dropdown "Điểm" tren thanh tren (neu co), cot Joint de trong de user
    /// chon trong bang. Trung diem thi bao va khong them. (Dropdown Joint
    //  ETABS don da xoa — joint chon truc tiep trong bang load point, user
    /// chot 08/11.)
    /// </summary>
    private void AddCurrentPointToPointGrid()
    {
        string point = _pointViewCombo.Text;
        if (_table is null)
        {
            MessageBox.Show(this, "Import file trước.", "Chưa có bảng",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        RefreshPointGrid();
        if (point != "" && _table.Points.Contains(point))
        {
            foreach (var (pt, _) in PointGridMappings())
            {
                if (pt == point)
                {
                    MessageBox.Show(this, $"Điểm '{point}' đã có trong bảng load point rồi.",
                        "Trùng điểm", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ShowPointGrid();
                    return;
                }
            }
            _pointGrid.Rows.Add(point, "");
        }
        else
        {
            _pointGrid.Rows.Add("", "");
        }
        ShowPointGrid();
    }

    /// <summary>
    /// Lam trang bang load point (user chot 08/11 — de gan luot moi): xoa
    /// het dong da nhap trong bang. Co xac nhan vi la thao tac pha huy.
    /// </summary>
    private void ClearPointGrid()
    {
        if (PointGridMappings().Count == 0 && _pointGrid.Rows.Count <= 1)
        {
            ShowPointGrid();
            return; // Bang da trang.
        }
        if (MessageBox.Show(this, "Xoá hết các điểm trong bảng load point để gán lượt mới?",
                "Làm trắng bảng", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;
        _pointGrid.Rows.Clear();
        ShowPointGrid();
    }

    /// <summary>
    /// Xoa dong dang chon trong bang load point (user chot 08/11 — de xoa
    /// dong thua: dong trong, dong nhap nham). Chon 1 hoac nhieu dong roi
    /// bam nut. Khong xoa dong nhap moi (dong trong cuoi bang).
    /// </summary>
    private void DeleteSelectedPointRows()
    {
        ShowPointGrid();
        var toDelete = new List<DataGridViewRow>();
        foreach (DataGridViewRow row in _pointGrid.SelectedRows)
        {
            if (!row.IsNewRow)
                toDelete.Add(row);
        }
        // Neu chi chon o (khong chon ca dong) thi lay dong cua o dang chon.
        if (toDelete.Count == 0 && _pointGrid.CurrentRow is not null && !_pointGrid.CurrentRow.IsNewRow)
            toDelete.Add(_pointGrid.CurrentRow);
        if (toDelete.Count == 0)
        {
            MessageBox.Show(this, "Chọn dòng cần xoá trong bảng load point trước (bấm vào đầu dòng để chọn).",
                "Chưa chọn dòng", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        foreach (var row in toDelete)
            _pointGrid.Rows.Remove(row);
        ShowPointGrid();
    }

    /// <summary>
    /// Do rong toi thieu cho 2 cot Pattern: du de thay het ten pattern dai
    /// nhat (trong model ETABS + profile), cong them cho nut dropdown.
    /// Dung lam MinimumWidth — cot van gian theo Fill khi cua so rong.
    /// </summary>
    private int PatternColumnWidth(IReadOnlyList<string> names)
    {
        int maxW = 0;
        foreach (string n in names)
        {
            if (n == "")
                continue;
            int w = TextRenderer.MeasureText(n, _grid.Font).Width;
            if (w > maxW)
                maxW = w;
        }
        return Math.Max(150, maxW + 40); // 40 = nut dropdown + padding.
    }

    private GroupBox BuildSectionAxis()
    {
        // Phuong & chieu gon 1 hang ngang (chi set 1 lan): Hx -> [X] [-]  Hy -> [Y] [+]  Vz -> [Z] [-].
        // Cao 80 (user chot 08/11 — tang them de chu khoi bi sat).
        var box = new GroupBox { Text = "Phương & chiều", Dock = DockStyle.Top, Height = 80 };
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
            // Doi phuong-chieu: chi cap nhat profile trong bo nho; luu bang
            // nut "Lưu data và Excel" (user chot 08/11 — khong tu dong luu).
            tgt.SelectedIndexChanged += (_, _) => { };
            sign.SelectedIndexChanged += (_, _) => { };
            x += 28; // khoang cach giua 2 anh xa
        }
        // Them nguoc (note truoc, row sau) de row o tren, note o duoi.
        box.Controls.Add(note);
        box.Controls.Add(row);
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
        // Pattern: doc tu 2 cot Pattern trong grid (neu grid dang hien bang);
        // neu grid chua co bang thi giu PatternMap hien tai (tranh xoa match
        // khi doi diem hoac chua chon diem).
        if (_table is not null && _grid.Columns.Contains("Pat1") && _grid.Rows.Count == _table.Rows.Count)
        {
            var map = new Dictionary<string, RowPattern>();
            for (int i = 0; i < _table.Rows.Count; i++)
            {
                string left = _grid.Rows[i].Cells["Pat1"].Value?.ToString() ?? "";
                string right = _grid.Rows[i].Cells["Pat2"].Value?.ToString() ?? "";
                map[_table.Rows[i].Name] = new RowPattern { Left = left, Right = right };
            }
            _profile.PatternMap = map;
        }
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
            RefreshGrid(); // Ve lai grid — 2 cot Pattern lay tu profile vua nap.
    }

    /// <summary>
    /// Luu THU CONG thong tin match vao sheet "LoadAssigner" trong file Excel
    /// noi luc (user chot 08/11 — nut "Lưu data và Excel", KHONG tu dong
    /// luu nua). Ghi TAT CA dong tai trong bang (ten tai + pattern 2 cot) +
    /// phuong-chieu + ghi de. Chi luu khi bang da nhan dien xong (co dong
    /// tai); bang rong thi bao loi va KHONG ghi — tranh ghi de sheet thanh
    /// trang nhu truoc day (tu dong luu khi sheet README 0 dong).
    /// File dang mo trong Excel thi khong ghi duoc — bao o thanh trang thai.
    /// </summary>
    private void SaveMatchToExcel()
    {
        if (!_uiReady || _table is null || _table.Rows.Count == 0)
        {
            MessageBox.Show(this,
                "Chưa có bảng tải để lưu.\nChọn file Excel rồi chọn đúng sheet có bảng tải (vd SH6-T3a) trước khi bấm Lưu.",
                "Chưa có bảng", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        string path = _profile.SourceFile;
        if (path == "" || !File.Exists(path))
        {
            MessageBox.Show(this, "Chưa chọn file Excel.", "Chưa chọn file",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        try
        {
            SyncProfileFromUi();
            Detector.WriteProfileSheet(path, _profile, _table);
            _statusLabel.Text = $"Đã lưu match vào sheet 'LoadAssigner' trong file Excel ({_table.Rows.Count} dòng tải, sheet: {_profile.Sheet}).";
        }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Không lưu được vào sheet 'LoadAssigner' trong file Excel ({ex.Message}) — đóng file trong Excel rồi bấm Lưu lại.";
            MessageBox.Show(this, _statusLabel.Text, "Không lưu được",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    /// <summary>
    /// Mo lai file Excel mo lan cuoi (nut "File gần nhất"). Neu file da bi
    /// xoa/doi ten thi bao user chon lai bang "Chọn file Excel...".
    /// </summary>
    private void OpenRecentFile()
    {
        string? recent = LoadRecentFile();
        if (recent is null)
        {
            MessageBox.Show(this, "Chưa có file Excel nào mở lần cuối.\nChọn file bằng nút “Chọn file Excel...”.",
                "Chưa có file gần nhất", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        _fileTextBox.Text = recent;
        _updatingSheet = true;
        try
        {
            _sheetCombo.Items.Clear();
            _sheetCombo.Items.AddRange(Detector.ListSheets(recent).ToArray());
            if (_sheetCombo.Items.Count > 0)
                _sheetCombo.SelectedIndex = 0;
        }
        catch { _sheetCombo.Items.Clear(); }
        finally { _updatingSheet = false; }
        Detect();
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
        SaveRecentFile(dlg.FileName); // Luu lam file mo lan cuoi (nut chon nhanh).
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
            // Fallback (user chot 08/11): sheet dang chon khong co data (vd
            // README — sheet dau tien khi mo file) thi tu dong chuyen sang
            // sheet tiep theo co data — DetectFile(path, null) quet tat ca
            // sheet va chon bang co nhieu dong tai nhat. Chi khi CA FILE
            // khong co data moi hien hop thoai bao loi (o cuoi ham, qua
            // t.Warnings).
            if (t.Rows.Count == 0 || t.Points.Count == 0)
            {
                var best = Detector.DetectFile(path, null);
                if (best.Rows.Count > 0 && best.Points.Count > 0)
                    t = best;
            }
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
        // Cap nhat dropdown Sheet theo sheet THUC TE da doc (neu fallback doi
        // sheet, vd README -> SH4-T1, thi dropdown hien SH4-T1 cho user biet).
        _updatingSheet = true;
        if (t.Sheet != "")
            _sheetCombo.Text = t.Sheet;
        _updatingSheet = false;
        _statusLabel.Text = $"Đã nhận diện: {t.Points.Count} điểm, {t.Rows.Count} dòng tải, {t.Patterns.Count} pattern (sheet: {t.Sheet}).";
        // Doc thong tin match tu sheet "LoadAssigner" trong chinh file Excel
        // nay (user chot 08/11 — thay file .json rieng). Co sheet thi nap
        // match/phuong-chieu/ghi de tu do; khong co thi dung mac dinh (goi y
        // theo ten) va se ghi sheet moi khi user doi match.
        try
        {
            var saved = Detector.ReadProfileSheet(path);
            if (saved is not null)
            {
                string src = _profile.SourceFile;
                string sh = _profile.Sheet;
                _profile = saved;
                _profile.SourceFile = src;
                _profile.Sheet = sh;
                ApplyProfileToUi();
                _statusLabel.Text += " Đã tải match từ sheet 'LoadAssigner' trong file.";
            }
        }
        catch (Exception ex)
        {
            _statusLabel.Text += $" (Không đọc được sheet 'LoadAssigner': {ex.Message} — dùng mặc định.)";
        }
        // Dropdown xem tung diem (ngay sau dropdown Sheet o thanh tren).
        // Chi hien 1 diem khi da chon (user chot 08/11 — khong load tat ca
        // diem de tranh load nhieu): mac dinh chua chon, grid de trong.
        // Diem gan = chinh diem nay (khong co dropdown chon diem rieng).
        _pointViewCombo.Items.Clear();
        _pointViewCombo.Items.AddRange(t.Points.ToArray());
        _pointViewCombo.SelectedIndex = -1;
        RefreshGrid();
        RefreshPointGrid(); // Cap nhat cot Diem Excel o bang load point.
        // Box thong tin file da xoa (user chot 08/11) — canh bao nhan dien
        // hien bang MessageBox de user thay ro.
        if (t.Warnings.Count > 0)
            MessageBox.Show(this, string.Join("\n", t.Warnings.Select(w => "  ! " + w)),
                "Cảnh báo nhận diện", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    // Ve lai grid theo diem dang chon o dropdown "Điểm" (ngay sau dropdown
    // Sheet): CHI HIEN 1 DIEM, kem tieu de ghi ro ten diem. Chua chon diem
    // -> grid de trong (user chot 08/11 — khong load tat ca diem de tranh
    // load nhieu). Ten diem chi o TIEU DE tren bang, cot la
    // Tải | Pattern giá trị 1 | Pattern giá trị 2 | Hx [kN] | Hy [kN] |
    // Vz [kN] (2 cot Pattern ngay sau cot Tai — user chot 08/11).
    private void RefreshGrid()
    {
        if (_table is null)
        {
            _grid.Columns.Clear();
            _grid.Rows.Clear();
            _pointTitleLabel.Text = "Chưa nhận diện bảng — chọn file Excel để tự nhận diện bảng.";
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
            // Cac cot Fill het chieu rong bang (user chot 08/11 — khong de
            // khoang trong xam ben phai). Ty trong: Tai 28, Pattern 1 = 26,
            // Pattern 2 = 26 (bang nhau), Hx/Hy/Vz 7 moi cot (rong hon truoc
            // de thay ro gia tri). 2 cot Pattern co MinimumWidth = ten
            // pattern dai nhat de khong bi cat khi cua so hep.
            var loadCol = new DataGridViewTextBoxColumn { Name = "Load", HeaderText = "Tải", ReadOnly = true, FillWeight = 28, MinimumWidth = 200 };
            _grid.Columns.Add(loadCol);
            // 2 cot Match pattern ngay sau cot Tai (user chot 08/11).
            var choices = PatternChoices();
            int patW = PatternColumnWidth(choices);
            var pat1 = new DataGridViewComboBoxColumn
            {
                Name = "Pat1", HeaderText = "Pattern — giá trị 1 (trái)",
                DataPropertyName = "Pat1", DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox,
                FillWeight = 26, MinimumWidth = patW,
            };
            pat1.Items.AddRange(choices.ToArray());
            _grid.Columns.Add(pat1);
            var pat2 = new DataGridViewComboBoxColumn
            {
                Name = "Pat2", HeaderText = "Pattern — giá trị 2 (phải)",
                DataPropertyName = "Pat2", DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox,
                FillWeight = 26, MinimumWidth = patW,
            };
            pat2.Items.AddRange(choices.ToArray());
            _grid.Columns.Add(pat2);
            // Header 2 dong: Hx / [kN] (nhu anh mau "85 / E.2").
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Hx", HeaderText = "Hx\n[kN]", FillWeight = 7, MinimumWidth = 60 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Hy", HeaderText = "Hy\n[kN]", FillWeight = 7, MinimumWidth = 60 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Vz", HeaderText = "Vz\n[kN]", FillWeight = 7, MinimumWidth = 60 });
            // Cot mo men Mx/My/Mz (don vi kN·m): TU AN khi bang khong co mo men
            // (user chot 08/11 — mo rong cho tuong lai, bang hien tai chi co luc).
            // Visible=false thi DataGridView Fill khong chia khoang cho cot do.
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Mx", HeaderText = "Mx\n[kN·m]", FillWeight = 7, MinimumWidth = 60, Visible = t.HasMoments });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "My", HeaderText = "My\n[kN·m]", FillWeight = 7, MinimumWidth = 60, Visible = t.HasMoments });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Mz", HeaderText = "Mz\n[kN·m]", FillWeight = 7, MinimumWidth = 60, Visible = t.HasMoments });
            List<string> available;
            try { available = _etabs.PatternNames(); }
            catch { available = new List<string>(); }
            foreach (var r in t.Rows)
            {
                var rp = _profile.PatternFor(r.Name);
                string curL = rp.Left;
                string curR = rp.Right;
                if (curL == "")
                    curL = Mapping.SuggestPattern(r.Name, available);
                // Goi y cot phai khi dong co o 2 gia tri (khoang "a/b" hoac "+/-V") — cot "+/-" khong tinh.
                if (curR == "" && r.HasRangeCell() && curL != "")
                    curR = Mapping.SuggestPatternOpposite(curL, available);
                var vals = new List<object> { r.Name + (r.Reversible ? "  (+/-)" : ""), curL, curR };
                // Hx/Hy/Vz + Mx/My/Mz (mo men chi co gia tri khi bang co mo men;
                // cot mo men an thi gia tri van luu trong dong, khong hien).
                foreach (string comp in Detector.AllComponents)
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
        // Chi doi tieu de khi dang hien bang load pattern (khong ghi de tieu
        // de "Bảng load point" khi user dang o bang phu).
        if (!_pointGrid.Visible)
            _pointTitleLabel.Text = point;
        _statusLabel.Text = $"Đang xem điểm {point} — bảng {t.Sheet} ({t.Rows.Count} dòng tải). Sửa trực tiếp trong bảng; giá trị mới dùng cho Gán.";
    }

    /// <summary>
    /// Lua chon cho 2 cot Pattern trong grid: "" (khong gan) + pattern CO
    /// SAN trong model (sorted) + cac pattern dang co trong profile (ke ca
    /// khi khong con trong model — de khong loi DataError khi hien).
    /// Tool KHONG tu tao pattern (user chot 08/11).
    /// </summary>
    private List<string> PatternChoices()
    {
        var choices = new List<string> { "" };
        var set = new SortedSet<string>();
        try { foreach (string p in _etabs.PatternNames()) set.Add(p); } catch { }
        if (_table is not null)
            foreach (var r in _table.Rows)
            {
                var rp = _profile.PatternFor(r.Name);
                if (rp.Left != "") set.Add(rp.Left);
                if (rp.Right != "") set.Add(rp.Right);
            }
        choices.AddRange(set);
        return choices;
    }

    /// <summary>
    /// Nap lai lua chon cho 2 cot Pattern trong grid (khi pattern trong
    /// model ETABS thay doi). Khong doi gia tri cell dang co. Cap nhat ca
    /// MinimumWidth theo ten pattern dai nhat moi.
    /// </summary>
    private void RefreshPatternColumns()
    {
        if (!_grid.Columns.Contains("Pat1") || !_grid.Columns.Contains("Pat2"))
            return;
        var choices = PatternChoices();
        var c1 = (DataGridViewComboBoxColumn)_grid.Columns["Pat1"];
        var c2 = (DataGridViewComboBoxColumn)_grid.Columns["Pat2"];
        c1.Items.Clear();
        c1.Items.AddRange(choices.ToArray());
        c2.Items.Clear();
        c2.Items.AddRange(choices.ToArray());
        int patW = PatternColumnWidth(choices);
        c1.MinimumWidth = patW;
        c2.MinimumWidth = patW;
    }

    // Sua truc tiep trong grid:
    //  - Cot Pattern (Pat1/Pat2): ghi vao profile + tu luu profile mac dinh.
    //  - Cot luc (Hx/Hy/Vz): ghi nguoc vao FoundationTable (Gan doc tu bang
    //    nay) roi chuan hoa hien thi theo Raw da parse.
    // Ten cot la ten thanh phan ("Hx"/"Hy"/"Vz"); diem lay tu dropdown
    // "Điểm" tren thanh tren (grid chi hien 1 diem).
    private void Grid_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (_updatingGrid || _table is null || e.RowIndex < 0 || e.ColumnIndex <= 0)
            return;
        if (e.RowIndex >= _table.Rows.Count)
            return;
        string colName = _grid.Columns[e.ColumnIndex].Name;
        var row = _table.Rows[e.RowIndex];
        if (colName == "Pat1" || colName == "Pat2")
        {
            // Doi pattern trong grid: ghi vao profile trong bo nho (luu bang
            // nut "Lưu data và Excel" — user chot 08/11, khong tu dong luu).
            string val = _grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
            var rp = _profile.PatternFor(row.Name);
            if (colName == "Pat1") rp.Left = val; else rp.Right = val;
            _profile.PatternMap[row.Name] = rp;
            return;
        }
        string comp = colName;
        string point = _pointViewCombo.Text;
        if (point == "")
            return;
        string newRaw = _grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
        var lv = Detector.ParseCell(newRaw);
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

    private void UpdateEtabsStatus()
    {
        try
        {
            var pats = _etabs.PatternNames();
            var pts = _etabs.ListPoints();
            _statusLabel.Text = $"ETABS: {pats.Count} pattern, {pts.Count} point.";
            RefreshJointCombo();
            RefreshPatternColumns();
        }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Không lấy được thông tin ETABS: {ex.Message}";
        }
    }

    /// <summary>
    /// Nap bang Joint ETABS: chi joint DANG CHON trong ETABS (user chot
    /// <summary>
    /// Nap danh sach joint ETABS: chi joint DANG CHON trong ETABS (user chot
    /// 08/11 — khong phai toan bo joint cua model), xep theo Z (thap truoc
    /// — joint mong o duoi), roi X, Y. Danh sach nay nap vao cot Joint o
    /// bang load point (chon joint truc tiep trong bang — dropdown Joint don
    /// da xoa, user chot 08/11).
    /// </summary>
    private void RefreshJointCombo()
    {
        List<PointInfo> sel;
        try { sel = _etabs.GetSelectedPoints(); }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Không lấy được joint đang chọn từ ETABS: {ex.Message}";
            return;
        }
        _jointList = sel.OrderBy(p => p.Z).ThenBy(p => p.X).ThenBy(p => p.Y).ThenBy(p => p.Name).ToList();
        if (_jointList.Count == 0)
            _statusLabel.Text = "Chưa chọn joint nào trong ETABS — chọn joint trong model trước khi mở plugin (cột Joint ở bảng load point chỉ liệt kê joint đang chọn).";
        else
            RefreshPointGrid(); // Cap nhat cot Joint o bang load point.
    }

    /// <summary>
    /// GAN HANG LOAT (user chot 08/11 — bang load point): gan TAT CA cap
    /// (Diem Excel -> Joint ETABS) trong bang load point, moi diem dung
    /// chung profile match pattern + phuong/chieu. Bang load point rong thi
    /// bao user them diem (khong gan 1 diem qua dropdown nua).
    /// </summary>
    private void Assign()
    {
        SyncProfileFromUi();
        if (_table is null)
        {
            MessageBox.Show(this, "Import file trước.", "Chưa có bảng",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        var mappings = PointGridMappings();
        if (mappings.Count == 0)
        {
            ShowPointGrid();
            MessageBox.Show(this,
                "Chưa có điểm nào trong bảng load point.\nThêm điểm bằng cách: chọn Điểm + Joint ở 2 dropdown rồi bấm “Thêm điểm”, hoặc gõ trực tiếp trong bảng load point (chọn Điểm Excel và Joint ETABS ở từng dòng).",
                "Chưa có điểm để gán", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        // Trung diem Excel: chi gan lan xuat hien dau (lan sau se xoa trang
        // diem vua gan neu tick Ghi de).
        var seen = new HashSet<string>();
        var dupes = new List<string>();
        var unique = new List<(string ExcelPoint, string JointName)>();
        foreach (var m in mappings)
        {
            if (!seen.Add(m.ExcelPoint))
                dupes.Add(m.ExcelPoint);
            else
                unique.Add(m);
        }
        if (dupes.Count > 0)
            MessageBox.Show(this, "Điểm trùng trong bảng load point (chỉ gán lần đầu): " + string.Join(", ", dupes.Distinct()),
                "Điểm trùng", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        // Build tat ca assignment truoc (gom ca canh bao bo qua) de xac nhan 1 lan.
        var perPoint = new List<(string ExcelPoint, string JointName, BuildResult Result)>();
        var allPats = new HashSet<string>();
        int totalAssign = 0;
        foreach (var (pt, jt) in unique)
        {
            var res = Assigner.BuildAssignments(_table, pt, jt, _profile);
            perPoint.Add((pt, jt, res));
            totalAssign += res.Assignments.Count;
            foreach (var a in res.Assignments)
                allPats.Add(a.Pattern);
        }
        if (totalAssign == 0)
        {
            MessageBox.Show(this, "Không có gì để gán: chưa match pattern hoặc toàn dòng 0.",
                "Không có gì để gán", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        var pats = allPats.OrderBy(x => x).ToList();
        string clearNote = _profile.ReplaceExisting
            ? "\n(Xoá trắng mọi tải cũ tại mỗi joint trước khi gán — điểm chỉ còn tải do plugin nhập vào.)"
            : "";
        string pointList = string.Join("\n", unique.Select(m => $"  • {m.ExcelPoint}  ->  {m.JointName}"));
        if (MessageBox.Show(this,
                $"Gán {totalAssign} tải cho {unique.Count} điểm?\n{pointList}\nPattern: {string.Join(", ", pats)}{clearNote}",
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
        var modelJointNames = new HashSet<string>(pts.Select(p => p.Name));
        var missingJoints = unique.Select(m => m.JointName).Where(j => !modelJointNames.Contains(j)).Distinct().ToList();
        if (missingJoints.Count > 0)
        {
            MessageBox.Show(this,
                "Joint không có trong model ETABS đang mở: " + string.Join(", ", missingJoints) +
                "\nChọn lại joint trong ETABS rồi gán lại.",
                "Joint không tồn tại", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        // Gan tung diem: ghi de (user chot 08/11) XOA TRANG moi tai cu tai joint
        // truoc khi gan moi — ca tai khong co trong file Excel cung bi xoa, de
        // diem chi con tai do plugin nhap vao. Khong tick thi chi ghi de/ghi
        // them tung pattern (tai cu khac pattern van con).
        var errs = new List<string>();
        var done = new List<(string ExcelPoint, string JointName, int Count, int Cleared)>();
        foreach (var (pt, jt, res) in perPoint)
        {
            int cleared = 0;
            if (_profile.ReplaceExisting)
            {
                try { cleared = _etabs.ClearJointLoads(jt); }
                catch (Exception ex)
                {
                    errs.Add($"{pt} -> {jt}: không xoá được tải cũ: {ex.Message}");
                    continue;
                }
            }
            int ok = 0;
            foreach (var assignment in res.Assignments)
            {
                try { _etabs.AssignForce(assignment.EtabsJoint, assignment.Pattern, assignment.Values(), _profile.ReplaceExisting); ok++; }
                catch (Exception ex) { errs.Add($"{assignment.Pattern}@{assignment.EtabsJoint} ({pt}): {ex.Message}"); }
            }
            done.Add((pt, jt, ok, cleared));
        }
        try { _etabs.RefreshView(); } catch { /* bo qua */ }
        var summary = new List<string>();
        foreach (var (pt, jt, count, cleared) in done)
            summary.Add($"  • {pt} -> {jt}: {count} tải" + (_profile.ReplaceExisting ? $" (đã xoá trắng {cleared} tải cũ)" : ""));
        if (errs.Count > 0)
            MessageBox.Show(this, "Đã gán:\n" + string.Join("\n", summary) + "\n\nLỗi:\n" + string.Join("\n", errs),
                "Gán lỗi một phần", MessageBoxButtons.OK, MessageBoxIcon.Error);
        else
            MessageBox.Show(this, $"Đã gán {totalAssign} tải cho {unique.Count} điểm:\n" + string.Join("\n", summary),
                "Xong", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

}
