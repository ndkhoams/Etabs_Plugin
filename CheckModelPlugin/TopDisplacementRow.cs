namespace Etabs_Ultimate_Tools
{
    public class TopDisplacementRow
    {
        public string Direction { get; set; }
        public string Combo { get; set; }
        public string TopStory { get; set; } // Story
        public double StoryElevation { get; set; } // m - cao độ tầng ETABS
        public double TopElevation { get; set; } // m - H tính từ mặt móng/ngàm đến tầng đang xét
        public double TopDisplacement { get; set; } // m
        public double TopDisplacementMm { get { return TopDisplacement * 1000.0; } }
        public double Ratio { get; set; } // Delta/H
        public double RatioDenominator { get { return Ratio > 1e-12 ? 1.0 / Ratio : 0.0; } }
        public double LimitDenominator { get; set; }
        public string Check { get; set; }
    }
}
