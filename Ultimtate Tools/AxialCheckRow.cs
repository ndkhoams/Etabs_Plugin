namespace Etabs_Ultimate_Tools
{
    /// <summary>
    /// Một hàng kết quả kiểm tra lực dọc quy đổi (cột hoặc vách Pier).
    /// </summary>
    public sealed class AxialCheckRow
    {
        public int    STT         { get; set; }
        public string Story       { get; set; } = "";
        public string Element     { get; set; } = "";
        public string ElementType { get; set; } = ""; // "Column" | "Pier.Col" | "Pier.Wall"
        public string Combo       { get; set; } = "";
        public string Material    { get; set; } = "";

        // Bê tông
        public double FckCube { get; set; }
        public double Fck     { get; set; }
        public double Fcd     { get; set; }

        // Nội lực & hình học
        public double Ned   { get; set; }
        public double T3    { get; set; } // Cột: chiều cao t3 | Vách: bề dày
        public double T2    { get; set; } // Cột: bề rộng t2  | Vách: chiều dài tổng
        public double Ac    { get; set; }
        public double AcFcd { get; set; }

        // Kết quả kiểm tra
        public double NuD     { get; set; }
        public double VdLimit { get; set; }
        public string Result  { get; set; } = "";
    }
}
