namespace Etabs_Ultimate_Tools
{
    /// <summary>Một dòng kết quả kiểm tra chuyển vị lệch tầng do gió.</summary>
    public class WindDriftRow
    {
        public string Direction { get; set; }
        public string Combo { get; set; }
        public string Story { get; set; }
        public double Elevation { get; set; }
        public double Height { get; set; }
        public double Drift { get; set; }                 // tỉ số Δ/h từ ETABS Story Drifts
        public double DriftDisplacement { get; set; }     // Δ = drift × h (m)
        public double LimitDenominator { get; set; }      // mẫu số giới hạn, vd 500
        public double DriftRatioDenominator => Drift > 1e-12 ? 1.0 / Drift : 0.0;
        public string Check { get; set; }
    }
}
