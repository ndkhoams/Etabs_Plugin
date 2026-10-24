using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    public partial class ModelCheckForm
    {
        private const string UpdateOwner = "ndkhoams";
        private const string UpdateRepository = "Etabs_Plugin";
        private const string UpdateFilePath = "Etabs_Tool.iso";
        private const string UpdateBranch = "main";
        private const string UpdateVersionFileName = "version.txt";
        private static readonly HttpClient UpdateHttpClient = CreateUpdateHttpClient();

        // Trạng thái của lần kiểm tra cập nhật gần nhất: chỉ bật nút Tải khi ISO trên
        // GitHub thật sự mới hơn bản đang chạy; giữ lại bytes ISO đã tải khi kiểm tra
        // để bước Tải ghi thẳng ra thay vì tải lại lần nữa.
        private bool _latestUpdateAvailable;
        private byte[] _checkedIsoBytes;

        private static HttpClient CreateUpdateHttpClient()
        {
            var client = new HttpClient { Timeout = TimeSpan.FromMinutes(3) };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("Etabs-Ultimate-Tools-Updater");
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
            return client;
        }

        private static string GetPluginBuildTimestamp()
        {
            var metadata = typeof(ModelCheckForm).Assembly
                .GetCustomAttributes(typeof(AssemblyMetadataAttribute), false)
                .Cast<AssemblyMetadataAttribute>()
                .FirstOrDefault(attribute => string.Equals(attribute.Key, "PluginBuildTimestamp",
                    StringComparison.OrdinalIgnoreCase));
            return metadata == null ? "00000000-000000" : metadata.Value;
        }

        private void BuildUpdateTab(TabPage tab)
        {
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 7,
                Padding = new Padding(24)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 150));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tab.Controls.Add(layout);

            var title = MakeTitle("CẬP NHẬT ETABS ULTIMATE TOOLS");
            title.Font = new System.Drawing.Font("Arial", 16F, System.Drawing.FontStyle.Bold);
            layout.Controls.Add(title, 0, 0);

            var versionBox = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "Phiên bản",
                Padding = new Padding(12),
                Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold)
            };
            layout.Controls.Add(versionBox, 0, 2);

            var versionLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2
            };
            versionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190));
            versionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            versionLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            versionLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            versionBox.Controls.Add(versionLayout);

            versionLayout.Controls.Add(MakeUpdateFieldLabel("Bản dựng hiện tại:"), 0, 0);
            lblUpdateBuildDate = MakeUpdateValueLabel(GetBuildDateDisplay());
            versionLayout.Controls.Add(lblUpdateBuildDate, 1, 0);
            versionLayout.Controls.Add(MakeUpdateFieldLabel("Bản dựng mới nhất:"), 0, 1);
            lblUpdateLatestDate = MakeUpdateValueLabel("Chưa kiểm tra");
            versionLayout.Controls.Add(lblUpdateLatestDate, 1, 1);

            lblUpdateStatus = new Label
            {
                Dock = DockStyle.Fill,
                Text = "Chưa kiểm tra phiên bản mới.",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(45, 55, 72)
            };
            layout.Controls.Add(lblUpdateStatus, 0, 3);

            progressUpdate = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Style = ProgressBarStyle.Blocks
            };
            layout.Controls.Add(progressUpdate, 0, 4);

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0)
            };
            btnUpdateCheck = MakeButton("Kiểm tra phiên bản");
            btnUpdateCheck.Width = 160;
            btnUpdateCheck.Font = new System.Drawing.Font("Arial", 10F);
            btnUpdateCheck.Click += async (s, e) => await CheckForPluginUpdateAsync();
            actions.Controls.Add(btnUpdateCheck);

            btnUpdateDownload = MakeButton("Tải file cập nhật");
            btnUpdateDownload.Width = 160;
            btnUpdateDownload.Font = new System.Drawing.Font("Arial", 10F);
            btnUpdateDownload.Enabled = false;
            btnUpdateDownload.Click += async (s, e) => await DownloadPluginUpdateAsync();
            actions.Controls.Add(btnUpdateDownload);
            layout.Controls.Add(actions, 0, 5);

            var note = MakeNote("File ISO sẽ được lưu tại vị trí bạn chọn. Tải file không tự cài đặt hoặc thay thế plugin đang chạy.");
            note.Font = new System.Drawing.Font("Arial", 11F);
            layout.Controls.Add(note, 0, 6);
        }

        private static Label MakeUpdateFieldLabel(string text)
        {
            return new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                Text = text,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Arial", 11F),
                Margin = new Padding(0)
            };
        }

        private static Label MakeUpdateValueLabel(string text)
        {
            return new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                Text = text,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Consolas", 11F),
                AutoEllipsis = true,
                Margin = new Padding(0)
            };
        }

        private string GetBuildDateDisplay()
        {
            DateTime buildDate;
            return TryGetBuildDate(out buildDate)
                ? buildDate.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)
                : "Không tìm thấy ngày build";
        }

        private bool TryGetBuildDate(out DateTime buildDate)
        {
            buildDate = default(DateTime);
            return DateTime.TryParseExact(currentBuild, "yyyyMMdd-HHmmss",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out buildDate);
        }

        private static string BuildIsoDownloadUrl(string commitSha)
        {
            string escapedPath = string.Join("/", UpdateFilePath.Split('/').Select(Uri.EscapeDataString));
            return "https://raw.githubusercontent.com/" + UpdateOwner + "/" + UpdateRepository +
                "/" + commitSha + "/" + escapedPath;
        }

        // Đọc dấu build trong version.txt đóng gói ở thư mục gốc của ISO.
        private static bool TryReadIsoBuildDate(byte[] isoBytes, out DateTime buildDate)
        {
            buildDate = default(DateTime);
            string content;
            if (!TryReadIsoTextFile(isoBytes, UpdateVersionFileName, out content))
                return false;
            string firstLine = (content ?? "").Trim().TrimStart('\uFEFF');
            int nl = firstLine.IndexOfAny(new[] { '\r', '\n' });
            if (nl >= 0) firstLine = firstLine.Substring(0, nl).Trim();
            return DateTime.TryParseExact(firstLine, "yyyyMMdd-HHmmss",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out buildDate);
        }

        // ── Đọc một file nhỏ ở thư mục gốc của ảnh ISO 9660 (kèm Joliet) ──────────
        // Vừa đủ để lấy version.txt mà không cần mount ISO hay dùng thư viện ngoài.
        private static bool TryReadIsoTextFile(byte[] iso, string fileName, out string content)
        {
            content = null;
            if (iso == null || iso.Length < 17 * 2048) return false;
            try
            {
                for (int sector = 16; sector < 32; sector++)
                {
                    int off = sector * 2048;
                    if (off + 2048 > iso.Length) break;
                    byte type = iso[off];
                    if (type == 255) break; // terminator của dãy volume descriptor
                    if (type != 1 && type != 2) continue;
                    if (iso[off + 1] != 'C' || iso[off + 2] != 'D' || iso[off + 3] != '0' ||
                        iso[off + 4] != '0' || iso[off + 5] != '1') continue;
                    bool joliet = type == 2;
                    // Bản ghi thư mục gốc nằm tại offset 156 trong volume descriptor.
                    int rootExtent = ReadLe32(iso, off + 156 + 2);
                    int rootSize = ReadLe32(iso, off + 156 + 10);
                    int fileExtent, fileSize;
                    if (FindIsoRootFile(iso, rootExtent, rootSize, fileName, joliet,
                            out fileExtent, out fileSize))
                    {
                        long start = (long)fileExtent * 2048;
                        if (start < 0 || start + fileSize > iso.Length) return false;
                        content = Encoding.UTF8.GetString(iso, (int)start, fileSize);
                        return true;
                    }
                }
            }
            catch { return false; }
            return false;
        }

        private static bool FindIsoRootFile(byte[] iso, int extent, int size, string fileName,
            bool joliet, out int fileExtent, out int fileSize)
        {
            fileExtent = 0; fileSize = 0;
            long baseOff = (long)extent * 2048;
            if (baseOff < 0 || baseOff >= iso.Length) return false;
            int dirLen = (int)Math.Min((long)size, iso.Length - baseOff);
            int pos = 0;
            while (pos < dirLen)
            {
                int recLen = iso[(int)(baseOff + pos)];
                if (recLen == 0)
                {
                    // Padding ở cuối sector: nhảy sang sector kế tiếp.
                    pos = ((pos / 2048) + 1) * 2048;
                    continue;
                }
                if (pos + recLen > dirLen) break;
                byte flags = iso[(int)(baseOff + pos + 25)];
                int nameLen = iso[(int)(baseOff + pos + 32)];
                if ((flags & 2) == 0 && nameLen > 0) // không phải thư mục
                {
                    string name;
                    if (joliet)
                    {
                        name = Encoding.BigEndianUnicode.GetString(
                            iso, (int)(baseOff + pos + 33), nameLen);
                    }
                    else
                    {
                        name = Encoding.ASCII.GetString(iso, (int)(baseOff + pos + 33), nameLen);
                        int semi = name.IndexOf(';');
                        if (semi >= 0) name = name.Substring(0, semi);
                    }
                    if (string.Equals((name ?? "").Trim(), fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        fileExtent = ReadLe32(iso, (int)(baseOff + pos + 2));
                        fileSize = ReadLe32(iso, (int)(baseOff + pos + 10));
                        return fileSize >= 0;
                    }
                }
                pos += recLen;
            }
            return false;
        }

        private static int ReadLe32(byte[] data, int offset)
        {
            return data[offset] | (data[offset + 1] << 8) |
                (data[offset + 2] << 16) | (data[offset + 3] << 24);
        }

        private async Task CheckForPluginUpdateAsync()
        {
            btnUpdateCheck.Enabled = false;
            btnUpdateDownload.Enabled = false;
            progressUpdate.Style = ProgressBarStyle.Marquee;
            lblUpdateStatus.Text = "Đang kiểm tra GitHub...";

            try
            {
                string url = "https://api.github.com/repos/" + UpdateOwner + "/" + UpdateRepository +
                    "/commits?path=" + Uri.EscapeDataString(UpdateFilePath) + "&per_page=1";
                using (var response = await UpdateHttpClient.GetAsync(url))
                {
                    response.EnsureSuccessStatusCode();
                    string json = await response.Content.ReadAsStringAsync();
                    var serializer = new DataContractJsonSerializer(typeof(GitHubCommit[]));
                    GitHubCommit[] commits;
                    using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                        commits = serializer.ReadObject(stream) as GitHubCommit[];

                    if (commits == null || commits.Length == 0 || string.IsNullOrWhiteSpace(commits[0].Sha))
                        throw new InvalidOperationException("GitHub không trả về commit cho file cập nhật.");

                    _latestUpdateCommit = commits[0].Sha;
                    _latestUpdateAvailable = false;
                    _checkedIsoBytes = null;

                    // Nguồn sự thật duy nhất về "phiên bản mới nhất" là CHÍNH FILE ISO sẽ
                    // được tải về (trước đây lấy currentBuild từ source trên main — hai
                    // nguồn khác nhau, có thể lệch nhau theo cả hai chiều). Tải ISO tại
                    // đúng commit mới nhất rồi đọc version.txt đóng gói bên trong nó.
                    byte[] isoBytes = await UpdateHttpClient.GetByteArrayAsync(
                        BuildIsoDownloadUrl(_latestUpdateCommit));
                    _checkedIsoBytes = isoBytes;

                    DateTime latestBuildDate;
                    bool fromVersionFile = TryReadIsoBuildDate(isoBytes, out latestBuildDate);
                    if (!fromVersionFile)
                    {
                        // ISO đời cũ chưa có version.txt: lùi về ngày commit của chính
                        // file ISO — vẫn cùng một nguồn (file ISO), chỉ kém chính xác hơn.
                        string commitDateText = commits[0].Commit != null && commits[0].Commit.Committer != null
                            ? commits[0].Commit.Committer.Date : null;
                        DateTimeOffset commitDate;
                        if (string.IsNullOrWhiteSpace(commitDateText) ||
                            !DateTimeOffset.TryParse(commitDateText, CultureInfo.InvariantCulture,
                                DateTimeStyles.RoundtripKind, out commitDate))
                            throw new InvalidOperationException(
                                "ISO không có version.txt và GitHub không trả về ngày commit của file ISO.");
                        latestBuildDate = commitDate.ToOffset(TimeSpan.FromHours(7)).DateTime;
                    }

                    var latestBuildTimestampGmt7 = new DateTimeOffset(latestBuildDate, TimeSpan.FromHours(7));
                    lblUpdateLatestDate.Text = latestBuildTimestampGmt7.ToString(
                        "yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
                    DateTime buildDate;
                    if (!TryGetBuildDate(out buildDate))
                    {
                        lblUpdateBuildDate.Text = "Không tìm thấy ngày build";
                        lblUpdateStatus.Text = "Không đọc được ngày build (định dạng yyyyMMdd-HHmmss).";
                    }
                    else
                    {
                        lblUpdateBuildDate.Text = buildDate.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
                        var buildTimestampGmt7 = new DateTimeOffset(buildDate, TimeSpan.FromHours(7));
                        _latestUpdateAvailable = latestBuildTimestampGmt7 > buildTimestampGmt7;
                        lblUpdateStatus.Text = _latestUpdateAvailable
                            ? "Có phiên bản mới."
                            : "Bạn đang dùng phiên bản mới nhất.";
                        if (!fromVersionFile)
                            lblUpdateStatus.Text += " (ISO chưa có version.txt — so theo ngày commit của file ISO.)";
                    }
                    btnUpdateDownload.Enabled = _latestUpdateAvailable;
                }
            }
            catch (Exception ex)
            {
                if (!IsDisposed)
                    lblUpdateStatus.Text = "Không kiểm tra được phiên bản: " + ex.Message;
            }
            finally
            {
                if (!IsDisposed)
                {
                    progressUpdate.Style = ProgressBarStyle.Blocks;
                    progressUpdate.Value = 0;
                    btnUpdateCheck.Enabled = true;
                }
            }
        }

        private async Task DownloadPluginUpdateAsync()
        {
            if (string.IsNullOrWhiteSpace(_latestUpdateCommit)) return;

            using (var dialog = new SaveFileDialog
            {
                Title = "Lưu file cập nhật",
                FileName = UpdateFilePath,
                DefaultExt = "iso",
                AddExtension = true,
                OverwritePrompt = true,
                Filter = "ISO image (*.iso)|*.iso|All files (*.*)|*.*"
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;

                string destination = dialog.FileName;
                string temporaryFile = Path.Combine(Path.GetDirectoryName(destination),
                    Path.GetFileName(destination) + "." + Guid.NewGuid().ToString("N") + ".download");
                btnUpdateCheck.Enabled = false;
                btnUpdateDownload.Enabled = false;
                progressUpdate.Value = 0;
                lblUpdateStatus.Text = "Đang tải commit " + _latestUpdateCommit + "...";

                try
                {
                    if (_checkedIsoBytes != null)
                    {
                        // Đã tải đúng ISO này ở bước kiểm tra — ghi thẳng ra, không tải lại.
                        File.WriteAllBytes(temporaryFile, _checkedIsoBytes);
                        progressUpdate.Style = ProgressBarStyle.Blocks;
                        progressUpdate.Value = 100;
                    }
                    else using (var response = await UpdateHttpClient.GetAsync(
                        BuildIsoDownloadUrl(_latestUpdateCommit), HttpCompletionOption.ResponseHeadersRead))
                    {
                        response.EnsureSuccessStatusCode();
                        long totalBytes = response.Content.Headers.ContentLength.GetValueOrDefault();
                        progressUpdate.Style = totalBytes > 0 ? ProgressBarStyle.Blocks : ProgressBarStyle.Marquee;

                        using (var input = await response.Content.ReadAsStreamAsync())
                        using (var output = new FileStream(temporaryFile, FileMode.CreateNew, FileAccess.Write,
                            FileShare.None, 81920, true))
                        {
                            byte[] buffer = new byte[81920];
                            long downloaded = 0;
                            int read;
                            while ((read = await input.ReadAsync(buffer, 0, buffer.Length)) > 0)
                            {
                                await output.WriteAsync(buffer, 0, read);
                                downloaded += read;
                                if (totalBytes > 0)
                                    progressUpdate.Value = Math.Min(100, (int)(downloaded * 100.0 / totalBytes));
                            }
                        }
                    }

                    if (File.Exists(destination))
                        File.Replace(temporaryFile, destination, null);
                    else
                        File.Move(temporaryFile, destination);

                    lblUpdateStatus.Text = "Đã tải xong commit " + _latestUpdateCommit + ".";
                    Info("Đã tải file cập nhật về:\n" + destination, "Update");
                }
                catch (Exception ex)
                {
                    if (!IsDisposed)
                        lblUpdateStatus.Text = "Tải file thất bại: " + ex.Message;
                }
                finally
                {
                    if (File.Exists(temporaryFile))
                    {
                        try { File.Delete(temporaryFile); }
                        catch { }
                    }
                    if (!IsDisposed)
                    {
                        progressUpdate.Style = ProgressBarStyle.Blocks;
                        progressUpdate.Value = 0;
                        btnUpdateCheck.Enabled = true;
                        btnUpdateDownload.Enabled = _latestUpdateAvailable;
                    }
                }
            }
        }

        [DataContract]
        private sealed class GitHubCommit
        {
            [DataMember(Name = "sha")]
            public string Sha { get; set; }

            [DataMember(Name = "commit")]
            public GitHubCommitMetadata Commit { get; set; }
        }

        [DataContract]
        private sealed class GitHubCommitMetadata
        {
            [DataMember(Name = "committer")]
            public GitHubCommitPerson Committer { get; set; }
        }

        [DataContract]
        private sealed class GitHubCommitPerson
        {
            [DataMember(Name = "date")]
            public string Date { get; set; }
        }
    }
}