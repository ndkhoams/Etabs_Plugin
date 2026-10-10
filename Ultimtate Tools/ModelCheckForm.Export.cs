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
        // ---------- Xuất Excel ----------

        private void RunExport(Action<string> writer, string suggestedName)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = suggestedName;
                if (sfd.ShowDialog() != DialogResult.OK) return;

                writer(sfd.FileName);
                Info("Đã xuất: " + sfd.FileName, "Xuất Excel");
            }
        }

        private void ExportPDelta()
        {
            if (_rows == null || _rows.Count == 0)
            {
                Warn("Chưa có dữ liệu P-Delta để xuất. Hãy bấm Tính kiểm tra trước.", "Xuất Excel");
                return;
            }
            RunExport(file => ExcelExporter.Export(file, _rows, _qFactor), "P-Delta.xlsx");
        }

        private void ExportWind()
        {
            if (_windRows == null || _windRows.Count == 0)
            {
                Warn("Chưa có dữ liệu chuyển vị đỉnh để xuất. Hãy bấm Tính kiểm tra trước.", "Xuất Excel");
                return;
            }
            RunExport(file => ExcelExporter.Export(file, null, _qFactor, _windRows), "ChuyenViDinh_Gio.xlsx");
        }

        private void ExportWindDrift()
        {
            if (_windDriftRows == null || _windDriftRows.Count == 0)
            {
                Warn("Chưa có dữ liệu chuyển vị lệch tầng do gió để xuất. Hãy bấm Tính kiểm tra trước.", "Xuất Excel");
                return;
            }
            RunExport(file => ExcelExporter.Export(file, null, _qFactor, null, _windDriftRows), "ChuyenViLechTang_Gio.xlsx");
        }

        private void ExportSeismic()
        {
            if (_seismicDriftRows == null || _seismicDriftRows.Count == 0)
            {
                Warn("Chưa có dữ liệu chuyển vị lệch tầng do động đất để xuất. Hãy bấm Tính kiểm tra trước.", "Xuất Excel");
                return;
            }
            RunExport(file => ExcelExporter.Export(file, null, _qFactor, null, null, _seismicDriftRows), "ChuyenViLechTang_DongDat.xlsx");
        }

        private void ExportAxial()
        {
            if (_axialRows == null || _axialRows.Count == 0)
            {
                Warn("Chưa có dữ liệu lực dọc để xuất. Hãy bấm Kiểm tra trước.", "Xuất Excel");
                return;
            }
            RunExport(file => AxialCheckExporter.Export(_axialRows, file, AxialAlphaCc, AxialGammaC, AxialColumnLimit, AxialWallLimit), "KiemTra_LucDoc.xlsx");
        }
    }
}
