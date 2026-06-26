using ETABSv1;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CheckModelPlugin
{
    public static class TopDisplacementExtractor
    {
        private class StoryInfo
        {
            public string Name;
            public double Elevation;
            public double Height;
        }

        public static List<TopDisplacementRow> Calculate(cSapModel sap, string comboX, string comboY, double limitDenominator)
        {
            var rows = new List<TopDisplacementRow>();
            var stories = ReadStories(sap);
            if (stories.Count == 0) return rows;

            // H được tính từ mặt móng/ngàm đến tầng đang xét.
            // Mặt móng/ngàm được xác định theo kết quả chuyển vị diaphragm: tầng có chuyển vị X/Y bằng 0
            // hoặc xấp xỉ 0 trong bảng Diaphragm Center of Mass Displacements.
            // Không lọc theo cao độ tầng > 0 nữa; chỉ cần H > 0 là đưa vào bảng kiểm tra.
            // Nhờ vậy mô hình có nhiều tầng hầm vẫn kiểm tra đúng các tầng phía trên mặt ngàm.
            double baseElevation = FindFoundationElevation(sap, stories, comboX, comboY);

            var checkStories = stories
                .Select(s => new { Story = s, H = Math.Abs(s.Elevation - baseElevation) })
                // Chỉ kiểm tra các tầng có H > 0 và bỏ level Base/phụ trợ của ETABS.
                // Base thường là mặt tham chiếu nằm dưới tầng ngàm thực tế, không phải tầng công trình cần kiểm tra.
                .Where(x => x.H > 1e-9 && !IsBaseLevel(x.Story.Name))
                .OrderByDescending(x => x.Story.Elevation)
                .ToList();

            if (!string.IsNullOrWhiteSpace(comboX))
            {
                foreach (var item in checkStories)
                {
                    var st = item.Story;
                    rows.Add(CalculateOne(sap, comboX, "X", st.Name, st.Elevation, item.H, limitDenominator, stories));
                }
            }
            if (!string.IsNullOrWhiteSpace(comboY))
            {
                foreach (var item in checkStories)
                {
                    var st = item.Story;
                    rows.Add(CalculateOne(sap, comboY, "Y", st.Name, st.Elevation, item.H, limitDenominator, stories));
                }
            }
            return rows;
        }

        private static bool IsBaseLevel(string storyName)
        {
            if (string.IsNullOrWhiteSpace(storyName)) return false;
            string s = storyName.Trim();
            return s.Equals("Base", StringComparison.OrdinalIgnoreCase) ||
                   s.IndexOf("Base", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static TopDisplacementRow CalculateOne(cSapModel sap, string combo, string dir, string topStory, double storyElevation, double h, double limitDenominator, List<StoryInfo> stories)
        {
            // Ưu tiên lấy chuyển vị đỉnh từ bảng Diaphragm Center of Mass Displacements.
            // Đây là đúng nguồn để kiểm tra chuyển vị đỉnh theo diaphragm/center of mass.
            double u = ReadTopDisplacementFromDiaphragmTables(sap, combo, dir, topStory);

            // Dự phòng: nếu ETABS/DLL không expose được bảng diaphragm thì thử các bảng Story Displacements.
            if (Math.Abs(u) < 1e-12)
                u = ReadTopDisplacementFromStoryTables(sap, combo, dir, topStory);

            double ratio = h > 1e-9 ? Math.Abs(u) / h : 0.0;
            double limit = limitDenominator > 0 ? 1.0 / limitDenominator : 0.0;

            return new TopDisplacementRow
            {
                Direction = dir,
                Combo = combo,
                TopStory = topStory,
                StoryElevation = storyElevation,
                TopElevation = h,
                TopDisplacement = Math.Abs(u),
                Ratio = ratio,
                LimitDenominator = limitDenominator,
                Check = limit > 0 && ratio > limit ? "NG" : "OK"
            };
        }

        private static List<StoryInfo> ReadStories(cSapModel sap)
        {
            int numberStories = 0;
            string[] storyNames = null;
            double[] storyElevations = null;
            double[] storyHeights = null;
            bool[] isMasterStory = null;
            string[] similarToStory = null;
            bool[] spliceAbove = null;
            double[] spliceHeight = null;
            sap.Story.GetStories(ref numberStories, ref storyNames, ref storyElevations, ref storyHeights,
                ref isMasterStory, ref similarToStory, ref spliceAbove, ref spliceHeight);

            var list = new List<StoryInfo>();
            for (int i = 0; i < numberStories; i++)
                list.Add(new StoryInfo
                {
                    Name = storyNames[i],
                    Elevation = storyElevations[i],
                    Height = storyHeights != null && i < storyHeights.Length ? storyHeights[i] : 0.0
                });
            return list;
        }

        private static double FindFoundationElevation(cSapModel sap, List<StoryInfo> stories, string comboX, string comboY)
        {
            const double zeroTol = 1e-9; // m

            var xMap = string.IsNullOrWhiteSpace(comboX)
                ? new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
                : ReadStoryDisplacementMapFromDiaphragmTables(sap, comboX, "X");

            var yMap = string.IsNullOrWhiteSpace(comboY)
                ? new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
                : ReadStoryDisplacementMapFromDiaphragmTables(sap, comboY, "Y");

            // Duyệt từ dưới lên và chọn tầng có dữ liệu displacement thật, đồng thời UX/UY bằng 0.
            // Nếu công trình có nhiều tầng hầm, tầng ngàm thực tế vẫn được nhận diện bằng điều kiện chuyển vị bằng 0.
            foreach (var st in stories.OrderBy(s => s.Elevation))
            {
                bool hasData = false;
                bool isFixed = true;

                if (xMap.TryGetValue(st.Name, out var ux))
                {
                    hasData = true;
                    if (Math.Abs(ux) > zeroTol) isFixed = false;
                }

                if (yMap.TryGetValue(st.Name, out var uy))
                {
                    hasData = true;
                    if (Math.Abs(uy) > zeroTol) isFixed = false;
                }

                if (hasData && isFixed)
                    return st.Elevation;
            }

            // Fallback 1: nếu không tìm thấy tầng chuyển vị bằng 0, lấy tầng thấp nhất có dữ liệu diaphragm.
            foreach (var st in stories.OrderBy(s => s.Elevation))
            {
                if (xMap.ContainsKey(st.Name) || yMap.ContainsKey(st.Name))
                    return st.Elevation;
            }

            // Fallback 2: trường hợp model/bảng chưa có dữ liệu diaphragm.
            return stories.Min(s => s.Elevation);
        }

        private static Dictionary<string, double> ReadStoryDisplacementMapFromDiaphragmTables(cSapModel sap, string combo, string dir)
        {
            string[] tableNames =
            {
                "Diaphragm Center of Mass Displacements",
                "Diaphragm Center Of Mass Displacements",
                "Diaphragm Centers of Mass Displacements",
                "Diaphragm Centers Of Mass Displacements",
                "Story Diaphragm Displacements",
                "Story Diaphragm Center of Mass Displacements",
                "Story/Diaphragm Displacements"
            };

            return ReadStoryDisplacementMapFromCandidateTables(sap, tableNames, combo, dir);
        }

        private static Dictionary<string, double> ReadStoryDisplacementMapFromCandidateTables(cSapModel sap, string[] tableNames, string combo, string dir)
        {
            var bestWithCase = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            var bestAnyCase = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            foreach (var tableName in tableNames)
            {
                List<Dictionary<string, string>> table;
                try
                {
                    table = EtabsTableReader.ReadTable(sap, tableName, combo);
                }
                catch
                {
                    continue;
                }

                foreach (var row in table)
                {
                    string story = EtabsTableReader.Get(row, "Story", "StoryName", "Story Name", "Level");
                    if (string.IsNullOrWhiteSpace(story)) continue;
                    story = story.Trim();

                    string outputCase = EtabsTableReader.Get(row, "Output Case", "OutputCase", "Load Case", "LoadCase", "Case", "Combo", "Combination");
                    double u = ReadDirectionalDisplacement(row, dir);

                    // Nếu field là mm thì giá trị thường > 5; đổi về m.
                    if (Math.Abs(u) > 5.0) u /= 1000.0;
                    u = Math.Abs(u);

                    if (!bestAnyCase.ContainsKey(story) || u > bestAnyCase[story])
                        bestAnyCase[story] = u;

                    if (IsSameOrBlank(outputCase, combo))
                    {
                        if (!bestWithCase.ContainsKey(story) || u > bestWithCase[story])
                            bestWithCase[story] = u;
                    }
                }

                if (bestWithCase.Count > 0)
                    return bestWithCase;
            }

            // Một số bản ETABS bị cache OutputCase; nếu chỉ đọc được một bảng đã lọc sẵn thì dùng dữ liệu có sẵn.
            return bestAnyCase;
        }

        private static double ReadTopDisplacementFromDiaphragmTables(cSapModel sap, string combo, string dir, string topStory)
        {
            // Các tên bảng ETABS phổ biến. Bản ETABS 22 thường dùng chữ "of" thường.
            string[] tableNames =
            {
                "Diaphragm Center of Mass Displacements",
                "Diaphragm Center Of Mass Displacements",
                "Diaphragm Centers of Mass Displacements",
                "Diaphragm Centers Of Mass Displacements",
                "Story Diaphragm Displacements",
                "Story Diaphragm Center of Mass Displacements",
                "Story/Diaphragm Displacements"
            };

            return ReadTopDisplacementFromCandidateTables(sap, tableNames, combo, dir, topStory, true);
        }

        private static double ReadTopDisplacementFromStoryTables(cSapModel sap, string combo, string dir, string topStory)
        {
            string[] tableNames =
            {
                "Story Displacements",
                "Story Max Over Avg Displacements"
            };

            return ReadTopDisplacementFromCandidateTables(sap, tableNames, combo, dir, topStory, false);
        }

        private static double ReadTopDisplacementFromCandidateTables(cSapModel sap, string[] tableNames, string combo, string dir, string topStory, bool preferDiaphragm)
        {
            double bestWithCase = 0.0;
            double bestAnyCase = 0.0;

            foreach (var tableName in tableNames)
            {
                List<Dictionary<string, string>> table;
                try
                {
                    table = EtabsTableReader.ReadTable(sap, tableName, combo);
                }
                catch
                {
                    continue;
                }

                foreach (var row in table)
                {
                    string story = EtabsTableReader.Get(row, "Story", "StoryName", "Story Name", "Level");
                    if (!string.IsNullOrWhiteSpace(story) &&
                        !string.Equals(story.Trim(), topStory.Trim(), StringComparison.OrdinalIgnoreCase))
                        continue;

                    string outputCase = EtabsTableReader.Get(row, "Output Case", "OutputCase", "Load Case", "LoadCase", "Case", "Combo", "Combination");
                    double u = ReadDirectionalDisplacement(row, dir);

                    // Nếu bảng không có Story field nhưng đang là bảng diaphragm, vẫn lấy giá trị lớn nhất.
                    // Nếu field là mm thì giá trị thường > 5; đổi về m.
                    if (Math.Abs(u) > 5.0) u /= 1000.0;
                    u = Math.Abs(u);

                    if (u > bestAnyCase) bestAnyCase = u;
                    if (IsSameOrBlank(outputCase, combo) && u > bestWithCase) bestWithCase = u;
                }

                if (bestWithCase > 0) return bestWithCase;
            }

            // Một số bản ETABS bị cache OutputCase; nếu chỉ đọc được một bảng đã lọc sẵn thì dùng giá trị lớn nhất.
            return bestAnyCase;
        }

        private static double ReadDirectionalDisplacement(Dictionary<string, string> row, string dir)
        {
            bool isX = dir.Equals("X", StringComparison.OrdinalIgnoreCase);
            if (isX)
            {
                return EtabsTableReader.GetDouble(row,
                    "UX", "Ux", "UX m", "Ux m", "UX mm", "Ux mm",
                    "U1", "U1 m", "U1 mm",
                    "X", "X m", "X mm",
                    "X-Displ", "X Displ", "Displ X", "Translation X",
                    "Global X", "GlobalX", "X Translation");
            }

            return EtabsTableReader.GetDouble(row,
                "UY", "Uy", "UY m", "Uy m", "UY mm", "Uy mm",
                "U2", "U2 m", "U2 mm",
                "Y", "Y m", "Y mm",
                "Y-Displ", "Y Displ", "Displ Y", "Translation Y",
                "Global Y", "GlobalY", "Y Translation");
        }

        private static bool IsSameOrBlank(string outputCase, string selectedName)
        {
            if (string.IsNullOrWhiteSpace(outputCase)) return true;
            if (string.Equals(outputCase.Trim(), selectedName.Trim(), StringComparison.OrdinalIgnoreCase)) return true;
            return outputCase.IndexOf(selectedName, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   selectedName.IndexOf(outputCase, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
