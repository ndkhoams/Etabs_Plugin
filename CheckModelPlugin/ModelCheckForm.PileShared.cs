using System;
using System.Drawing;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    // Helper UI dùng chung cho tab Pile Reactions (trước đây nằm trong ModelCheckForm.Pile.cs).
    public partial class ModelCheckForm
    {
        private static Label MakeHeadCell(string text) => new Label
        {
            Text = text, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter,
            BorderStyle = BorderStyle.FixedSingle, BackColor = SystemColors.ControlLight,
            Margin = new Padding(0), Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold)
        };

        private static Label MakePileLabel(string text) => new Label
        {
            Text = text, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
        };

        private static ComboBox MakePileCombo() => new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 4)
        };

        private DataGridView CreateEditableGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                BackgroundColor = SystemColors.ControlLightLight,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 8, 0, 0)
            };
        }

        private static DataGridViewTextBoxColumn MakeCapCol(int width, bool readOnly)
        {
            return new DataGridViewTextBoxColumn
            {
                Width = width, ReadOnly = readOnly,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
        }

        private static double ParseCap(object value)
        {
            if (value == null) return 0.0;
            string s = value.ToString().Trim().Replace(",", ".");
            double d;
            return double.TryParse(s, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out d) ? d : 0.0;
        }
    }
}
