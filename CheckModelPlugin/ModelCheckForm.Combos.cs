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
        // ---------- Tải danh sách tổ hợp ----------

        private void LoadCombos()
        {
            var combos = PDeltaExtractor.GetLoadCombinations(_sap);
            foreach (var cbo in new[] { cboCombo, cboWindCombo, cboWindDriftCombo, cboSeisCombo, cboAxialCombo })
            {
                cbo.Items.Clear();
                cbo.Items.AddRange(combos.Cast<object>().ToArray());
            }

            SelectByKeyword(cboCombo, "EQ-SRSS", "Vtot", "EQ", "DD", "DONGDAT", "RS", "SPEC", "E");
            SelectByKeyword(cboWindCombo, "ENV_SLS_W", "WX", "WY", "WINDX", "WINDY", "GIOX", "GIOY");
            SelectByKeyword(cboWindDriftCombo, "ENV_SLS_W", "WX", "WY", "WINDX", "WINDY", "GIOX", "GIOY");
            SelectByKeyword(cboSeisCombo, "EQ-SRSS", "Vtot", "DDX", "DDY", "DD", "DONGDAT", "RS", "SPEC", "E");

            // Tab Pile Reactions: mỗi trường hợp tải là 1 CheckedListBox cho chọn NHIỀU tổ hợp.
            foreach (var clb in new[] { clbPileHVert, clbPileHWind, clbPileHEq })
            {
                if (clb == null) continue;
                clb.Items.Clear();
                foreach (var name in combos)
                    clb.Items.Add(name, false);
            }
            // Tích sẵn theo yêu cầu: đứng = ULS01; gió = ULS02..ULS13; động đất = ULS14..ULS17.
            CheckUlsRange(clbPileHVert, 1, 1);
            CheckUlsRange(clbPileHWind, 2, 13);
            CheckUlsRange(clbPileHEq, 14, 17);
            LoadPileHSpringTypes();

            if (cboAxialCombo.Items.Count > 0 && cboAxialCombo.SelectedIndex < 0) cboAxialCombo.SelectedIndex = 0;

            if (clbColCombos != null)
            {
                clbColCombos.Items.Clear();
                foreach (var name in ColumnForceExporter.GetCombos(_sap))
                    clbColCombos.Items.Add(name, true);
            }
        }

        private static void SelectByKeyword(ComboBox cbo, params string[] keys)
        {
            foreach (var key in keys)
                for (int i = 0; i < cbo.Items.Count; i++)
                    if (string.Equals(cbo.Items[i].ToString(), key, StringComparison.OrdinalIgnoreCase)) { cbo.SelectedIndex = i; return; }

            foreach (var key in keys)
                for (int i = 0; i < cbo.Items.Count; i++)
                    if (cbo.Items[i].ToString().IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0) { cbo.SelectedIndex = i; return; }

            if (cbo.Items.Count > 0 && cbo.SelectedIndex < 0) cbo.SelectedIndex = 0;
        }

        // Tích sẵn các tổ hợp có chứa từ khóa (không phân biệt hoa thường) trong CheckedListBox.
        private static void CheckCombosByKeyword(CheckedListBox clb, params string[] keys)
        {
            if (clb == null) return;
            for (int i = 0; i < clb.Items.Count; i++)
            {
                string name = clb.Items[i].ToString();
                foreach (var key in keys)
                {
                    if (name.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        clb.SetItemChecked(i, true);
                        break;
                    }
                }
            }
        }

        // Tích sẵn các tổ hợp ULS có số thứ tự nằm trong [from, to] (vd ULS02..ULS13).
        private static void CheckUlsRange(CheckedListBox clb, int from, int to)
        {
            if (clb == null) return;
            for (int i = 0; i < clb.Items.Count; i++)
            {
                int num;
                if (TryGetUlsNumber(clb.Items[i].ToString(), out num) && num >= from && num <= to)
                    clb.SetItemChecked(i, true);
            }
        }

        // Lấy số thứ tự sau tiền tố "ULS" (ULS01 -> 1, ULS14 -> 14). Trả về false nếu không khớp.
        private static bool TryGetUlsNumber(string name, out int number)
        {
            number = 0;
            if (string.IsNullOrWhiteSpace(name)) return false;
            string trimmed = name.Trim();
            if (!trimmed.StartsWith("ULS", StringComparison.OrdinalIgnoreCase)) return false;
            string digits = new string(trimmed.Substring(3).TakeWhile(char.IsDigit).ToArray());
            return int.TryParse(digits, out number);
        }
    }
}
