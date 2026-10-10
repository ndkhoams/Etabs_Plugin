namespace Etabs_Ultimate_Tools
{
    /// <summary>Một dòng kết quả kiểm tra chuyển vị lệch tầng do động đất (TCVN 9386-1:2025).</summary>
    public class SeismicDriftRow
    {
        public string Direction { get; set; }
        public string Combo { get; set; }
        public string Story { get; set; }
        public double Elevation { get; set; }
        public double Height { get; set; }
        public double Q { get; set; }                  // hệ số ứng xử q
        public double Nu { get; set; }                 // hệ số chiết giảm ν
        public double Drift { get; set; }              // de/h - drift đàn hồi từ ETABS Story Drifts
        public double DesignDrift => Q * Drift;        // dr/h = q · de/h
        public double ReducedDrift => Q * Drift * Nu;  // (dr · ν)/h
        public double LimitRatio { get; set; }         // giới hạn 0.005 / 0.0075 / 0.010
        public double AllowableDrift => (Q * Nu) > 0 ? LimitRatio / (Q * Nu) : 0.0;  // ngưỡng so sánh: limit/(ν·q)
        public string Check { get; set; }
    }
}
