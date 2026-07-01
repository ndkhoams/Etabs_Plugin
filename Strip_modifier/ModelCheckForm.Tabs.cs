using ETABSv1;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    public partial class ModelCheckForm : Form
    {
        // ---------- Các tab ----------

        private void BuildPDeltaTab(TabPage tab)
        {
            dgv = BuildScaffold(tab,
                "KIỂM TRA ĐIỀU KIỆN P-DELTA",
                "(Theo TCVN 9386-1:2025)",
                "θ = dr / h × Ptot / Vtot =  q × drift × Ptot / Vtot (mục 4.4.2.2 eq. 4.28)",
                "drift = Δ/h được xác định từ hệ quả của tác động động đất thiết kế (mục 4.3.4); Vtot là tổng lực cắt tầng do động đất gây ra; 2 thành phần này đều lấy từ tổ hợp các thành phần phương ngang của động đất SRSS(EX;EY) (mục 4.3.3.5.1.2b). Ptot tự động lấy từ Mass Summary by Story.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp:", 68));
            cboCombo = MakeCombo(240); bar.Controls.Add(cboCombo);
            bar.Controls.Add(MakeFieldLabel("q:", 22));
            txtQ = MakeTextBox("1.5", 60); bar.Controls.Add(txtQ);

            btnRun = MakeButton("Tính toán"); btnRun.Click += (s, e) => RunCheck(); bar.Controls.Add(btnRun);
            btnExport = MakeButton("Xuất Excel"); btnExport.Enabled = false; btnExport.Click += (s, e) => ExportPDelta(); bar.Controls.Add(btnExport);
            btnClose = MakeButton("Đóng"); btnClose.Width = 84; btnClose.Click += (s, e) => Close(); bar.Controls.Add(btnClose);

            AddPDeltaGridColumns();
        }

        private void BuildWindTab(TabPage tab)
        {
            dgvWind = BuildScaffold(tab,
                "KIỂM TRA CHUYỂN VỊ ĐỈNH CÔNG TRÌNH",
                "(Theo TCVN 2737:2023)",
                "Điều kiện kiểm tra: f ≤ fu",
                "Giới hạn chuyển vị ngang tổng thể là H/500. H được tính là khoảng cách từ mặt móng đến mái.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp gió:", 78));
            cboWindCombo = MakeCombo(240); bar.Controls.Add(cboWindCombo);

            btnWindRun = MakeButton("Tính toán"); btnWindRun.Click += (s, e) => RunWindCheck(); bar.Controls.Add(btnWindRun);
            btnWindExport = MakeButton("Xuất Excel"); btnWindExport.Enabled = false; btnWindExport.Click += (s, e) => ExportWind(); bar.Controls.Add(btnWindExport);

            AddWindGridColumns();
        }

        private void BuildWindDriftTab(TabPage tab)
        {
            dgvWindDrift = BuildScaffold(tab,
                "KIỂM TRA CHUYỂN VỊ LỆCH TẦNG DO TẢI TRỌNG GIÓ",
                "(Theo TCVN 2737:2023)",
                "Điều kiện: drift = Δ/h ≤ 1/500 cho từng tầng",
                "Drift lấy trực tiếp từ ETABS Story Drifts theo tổ hợp gió.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp gió:", 78));
            cboWindDriftCombo = MakeCombo(240); bar.Controls.Add(cboWindDriftCombo);

            btnWindDriftRun = MakeButton("Tính toán"); btnWindDriftRun.Click += (s, e) => RunWindDriftCheck(); bar.Controls.Add(btnWindDriftRun);
            btnWindDriftExport = MakeButton("Xuất Excel"); btnWindDriftExport.Enabled = false; btnWindDriftExport.Click += (s, e) => ExportWindDrift(); bar.Controls.Add(btnWindDriftExport);

            AddWindDriftGridColumns();
        }

        private void BuildSeismicDriftTab(TabPage tab)
        {
            dgvSeis = BuildScaffold(tab,
                "KIỂM TRA CHUYỂN VỊ LỆCH TẦNG DO TẢI TRỌNG ĐỘNG ĐẤT",
                "(Theo TCVN 9386-1:2025)",
                "Điều kiện hạn chế hư hỏng: dr·ν ≤ limit·h  ⇔  drift ≤ limit/(ν·q) (mục 4.4.3.2)",
                "drift = de/h (đàn hồi) lấy từ ETABS Story Drifts. dr = q × de là chuyển vị ngang thiết kế tương đối giữa các tầng. Drift lấy từ tổ hợp các thành phần phương ngang của động đất SRSS(EX;EY)." +
                "\nCHÚ THÍCH: Các giá trị khác nhau của ν phụ thuộc vào các nguy cơ động đất và vào cấp hậu quả của công trình, khuyến nghị như sau: ν = 0,4 cho các cấp hậu quả C3-a và C3-b, và ν = 0,5 cho các cấp hậu quả C1 và C2.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp động đất:", 110));
            cboSeisCombo = MakeCombo(220); bar.Controls.Add(cboSeisCombo);
            bar.Controls.Add(MakeFieldLabel("q:", 22));
            txtSeisQ = MakeTextBox("1.5", 50); bar.Controls.Add(txtSeisQ);
            bar.Controls.Add(MakeFieldLabel("ν:", 22));
            txtSeisNu = MakeTextBox("0.4", 50); bar.Controls.Add(txtSeisNu);
            bar.Controls.Add(MakeFieldLabel("limit:", 38));
            cboSeisLimit = MakeCombo(150);
            cboSeisLimit.Items.AddRange(new object[] { "0.005 (giòn)", "0.0075 (dẻo)", "0.010 (không cản trở)" });
            cboSeisLimit.SelectedIndex = 0;
            bar.Controls.Add(cboSeisLimit);

            btnSeisRun = MakeButton("Tính toán"); btnSeisRun.Click += (s, e) => RunSeismicDriftCheck(); bar.Controls.Add(btnSeisRun);
            btnSeisExport = MakeButton("Xuất Excel"); btnSeisExport.Enabled = false; btnSeisExport.Click += (s, e) => ExportSeismic(); bar.Controls.Add(btnSeisExport);

            AddSeismicDriftGridColumns();
        }

        private void BuildAxialTab(TabPage tab)
        {
            dgvAxial = BuildScaffold(tab,
                "KIỂM TRA HỆ SỐ LỰC DỌC QUY ĐỔI",
                "(Theo TCVN 9386-1:2025)",
                "ʊd = Ned/(Ac·fcd) ≤ 0.65 (cột) / 0.40 (vách)  |  fcd = αcc·fck/γc",
                "Chọn cột hoặc vách (Pier) trong ETABS trước khi mở tool.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Bê tông:", 88));
            cboAxialConcrete = MakeCombo(110);
            cboAxialConcrete.Items.AddRange(new object[] { "B15", "B20", "B22.5", "B25", "B30", "B35", "B40", "B45", "B50", "B55", "B60", "B70", "B80" });
            cboAxialConcrete.SelectedItem = "B30";
            if (cboAxialConcrete.SelectedIndex < 0 && cboAxialConcrete.Items.Count > 0) cboAxialConcrete.SelectedIndex = 0;
            bar.Controls.Add(cboAxialConcrete);

            bar.Controls.Add(MakeFieldLabel("Combo:", 56));
            cboAxialCombo = MakeCombo(220); bar.Controls.Add(cboAxialCombo);

            btnAxialRun = MakeButton("Tính toán"); btnAxialRun.Click += (s, e) => RunAxialCheck(); bar.Controls.Add(btnAxialRun);
            btnAxialExport = MakeButton("Xuất Excel"); btnAxialExport.Enabled = false; btnAxialExport.Click += (s, e) => ExportAxial(); bar.Controls.Add(btnAxialExport);

            lblAxialInfo = new Label
            {
                AutoSize = false, Width = 280, Height = CtrlHeight,
                TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.DimGray,
                Margin = new Padding(8, 4, 0, 0)
            };
            bar.Controls.Add(lblAxialInfo);

            AddAxialGridColumns();
        }

        // ---------- Định nghĩa cột cho từng grid ----------

        private void AddPDeltaGridColumns()
        {
            dgv.Columns.Clear();
            AddColumn(dgv, "Direction", "Phương", 60);
            AddColumn(dgv, "Story", "Tầng", 80);
            AddColumn(dgv, "ElasticDrift", "drift", 120, "N5");
            AddColumn(dgv, "DesignDrift", "q × drift", 125, "N5");
            AddColumn(dgv, "Ptot", "Ptot (kN)", 105, "0");
            AddColumn(dgv, "Vtot", "Vtot (kN)", 105, "0");
            AddColumn(dgv, "Theta", "θ", 110, "N3");
            AddColumn(dgv, "Amplification", "1/(1-θ)", 90, "N3");
            AddColumn(dgv, "Conclusion", "Kết luận", 300, null, true);
        }

        private void AddWindGridColumns()
        {
            dgvWind.Columns.Clear();
            AddColumn(dgvWind, "Story", "Tầng", 90);
            AddColumn(dgvWind, "StoryElevation", "Cao độ tầng (m)", 120, "+0.000;-0.000;0.000");
            AddColumn(dgvWind, "Height", "H (m)", 110, "N3");
            AddColumn(dgvWind, "DeltaX", "ΔX (mm)", 120, "N1");
            AddColumn(dgvWind, "DeltaY", "ΔY (mm)", 120, "N1");
            AddColumn(dgvWind, "LimitMm", "H/500 (mm)", 120, "N0");
            AddColumn(dgvWind, "Check", "Kiểm tra", 250, null, true);
        }

        private void AddWindDriftGridColumns()
        {
            dgvWindDrift.Columns.Clear();
            AddColumn(dgvWindDrift, "Story", "Tầng", 90);
            AddColumn(dgvWindDrift, "Elevation", "Cao độ (m)", 100, "+0.000;-0.000;0.000");
            AddColumn(dgvWindDrift, "Height", "h tầng (m)", 90, "N3");
            AddColumn(dgvWindDrift, "DriftX", "drift X", 110, "0.000000");
            AddColumn(dgvWindDrift, "DriftY", "drift Y", 110, "0.000000");
            AddColumn(dgvWindDrift, "Limit", "Giới hạn 1/500", 120, "0.000000");
            AddColumn(dgvWindDrift, "Check", "Kiểm tra", 200, null, true);
        }

        private void AddSeismicDriftGridColumns()
        {
            dgvSeis.Columns.Clear();
            AddColumn(dgvSeis, "Story", "Tầng", 90);
            AddColumn(dgvSeis, "Elevation", "Cao độ (m)", 95, "+0.000;-0.000;0.000");
            AddColumn(dgvSeis, "Height", "h tầng (m)", 85, "N3");
            AddColumn(dgvSeis, "DriftX", "drift X (de/h)", 105, "0.000000");
            AddColumn(dgvSeis, "DriftY", "drift Y (de/h)", 105, "0.000000");
            AddColumn(dgvSeis, "DriftMax", "drift max", 100, "0.000000");
            AddColumn(dgvSeis, "AllowLimit", "Giới hạn limit/(ν·q)", 135, "0.000000");
            AddColumn(dgvSeis, "Check", "Kiểm tra", 150, null, true);
        }

        private void AddAxialGridColumns()
        {
            dgvAxial.Columns.Clear();
            AddColumn(dgvAxial, "STT", "STT", 45);
            AddColumn(dgvAxial, "Story", "Tầng", 80);
            AddColumn(dgvAxial, "ElementType", "Loại", 80);
            AddColumn(dgvAxial, "Element", "Label", 110);
            AddColumn(dgvAxial, "Combo", "Combo", 150);
            AddColumn(dgvAxial, "Ned", "Ned (kN)", 90, "0");
            AddColumn(dgvAxial, "T3", "t3 (m)", 70, "0.000");
            AddColumn(dgvAxial, "T2", "t2 (m)", 70, "0.000");
            AddColumn(dgvAxial, "Ac", "Ac (m²)", 80, "0.000");
            AddColumn(dgvAxial, "AcFcd", "Ac·fcd (kN)", 100, "0");
            AddColumn(dgvAxial, "NuD", "ʊd", 70, "0.000");
            AddColumn(dgvAxial, "VdLimit", "ʊd limit", 70, "0.00");
            AddColumn(dgvAxial, "Result", "Kết luận", 150, null, true);
        }
    }
}
