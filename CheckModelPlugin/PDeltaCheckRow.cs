namespace Etabs_Ultimate_Tools
{
    public class PDeltaCheckRow
    {
        public string Direction { get; set; }
        public string Story { get; set; }
        public double Elevation { get; set; }
        public double Height { get; set; }
        public double ElasticDrift { get; set; }
        public double DesignDrift { get; set; }
        public double Ptot { get; set; }
        public double Vtot { get; set; }
        public double Theta { get; set; }
        public double Amplification { get; set; }
        public string Conclusion { get; set; }
    }
}
