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
        private const string currentBuild = "20261010-085556";
        private static readonly HttpClient UpdateHttpClient = CreateUpdateHttpClient();

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
                    string latestDateText = commits[0].Commit?.Committer?.Date;
                    DateTimeOffset latestCommitDate;
                    if (!DateTimeOffset.TryParse(latestDateText, CultureInfo.InvariantCulture,
                        DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                        out latestCommitDate))
                        throw new InvalidOperationException("GitHub không trả về ngày commit hợp lệ.");

                    DateTimeOffset latestCommitDateGmt7 = latestCommitDate.ToOffset(TimeSpan.FromHours(7));
                    lblUpdateLatestDate.Text = latestCommitDateGmt7.ToString(
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
                        bool hasNewerVersion = latestCommitDateGmt7 > buildTimestampGmt7;
                        lblUpdateStatus.Text = hasNewerVersion
                            ? "Có phiên bản mới."
                            : "Bạn đang dùng phiên bản mới nhất.";
                    }
                    btnUpdateDownload.Enabled = true;
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
                    string escapedPath = string.Join("/", UpdateFilePath.Split('/').Select(Uri.EscapeDataString));
                    string url = "https://raw.githubusercontent.com/" + UpdateOwner + "/" + UpdateRepository +
                        "/" + _latestUpdateCommit + "/" + escapedPath;
                    using (var response = await UpdateHttpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead))
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
                        btnUpdateDownload.Enabled = !string.IsNullOrWhiteSpace(_latestUpdateCommit);
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