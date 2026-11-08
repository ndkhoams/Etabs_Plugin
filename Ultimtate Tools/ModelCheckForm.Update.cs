using System;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    public partial class ModelCheckForm
    {
        // Co che check for update giong CADtools (user chot 09/11): so theo
        // timestamp ghi tren title bar. BuildTimestamp la nguon su that duy
        // nhat — csproj tu dong dong dau moi lan build (target
        // UpdateBuildTimestamp), title bar hien dung gia tri nay, va buoc
        // kiem tra lay gia tri moi nhat tu chinh file nay tren GitHub.
        private const string BuildTimestamp = "20261109-004631";
        private const string UpdateOwner = "ndkhoams";
        private const string UpdateRepository = "Etabs_Plugin";
        private const string UpdateFilePath = "Etabs_Tool.iso";
        private const string LatestSourceUrl = "https://raw.githubusercontent.com/ndkhoams/Etabs_Plugin/main/Ultimtate%20Tools/ModelCheckForm.Update.cs";
        private const string DownloadUrl = "https://raw.githubusercontent.com/ndkhoams/Etabs_Plugin/main/Etabs_Tool.iso";
        private static readonly HttpClient UpdateHttpClient = CreateUpdateHttpClient();

        // Trang thai cua lan kiem tra cap nhat gan nhat: chi bat nut Tai khi
        // ban tren GitHub that su moi hon ban dang chay.
        private bool _latestUpdateAvailable;

        private static HttpClient CreateUpdateHttpClient()
        {
            var client = new HttpClient { Timeout = TimeSpan.FromMinutes(3) };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("Etabs-Ultimate-Tools-Updater");
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
            return client;
        }

        internal static DateTime BuildTimeLocal
        {
            get { return DateTime.ParseExact(BuildTimestamp, "yyyyMMdd-HHmmss", CultureInfo.InvariantCulture); }
        }

        private static async Task<DateTime> GetLatestBuildTimeAsync()
        {
            string source = await UpdateHttpClient.GetStringAsync(LatestSourceUrl).ConfigureAwait(false);
            Match match = Regex.Match(source, @"private const string BuildTimestamp = ""(?<timestamp>\d{8}-\d{6})"";");
            if (!match.Success)
                throw new InvalidDataException("Không tìm thấy timestamp build hợp lệ trong ModelCheckForm.Update.cs trên GitHub.");

            return DateTime.ParseExact(match.Groups["timestamp"].Value, "yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
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
            lblUpdateBuildDate = MakeUpdateValueLabel(BuildTimeLocal.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture));
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

        private async Task CheckForPluginUpdateAsync()
        {
            btnUpdateCheck.Enabled = false;
            btnUpdateDownload.Enabled = false;
            _latestUpdateAvailable = false;
            progressUpdate.Style = ProgressBarStyle.Marquee;
            lblUpdateStatus.Text = "Đang kiểm tra GitHub...";

            try
            {
                // Chi tai file .cs nho (~vai chuc KB) roi doc timestamp trong
                // do — khong tai ca file ISO nhu truoc, khong doc version.txt
                // trong ISO nua.
                DateTime latestBuildTime = await GetLatestBuildTimeAsync();
                if (IsDisposed) return;

                lblUpdateLatestDate.Text = latestBuildTime.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
                _latestUpdateAvailable = latestBuildTime > BuildTimeLocal;
                lblUpdateStatus.Text = _latestUpdateAvailable
                    ? "Có phiên bản mới."
                    : "Bạn đang dùng phiên bản mới nhất.";
                btnUpdateDownload.Enabled = _latestUpdateAvailable;
            }
            catch (Exception ex)
            {
                if (!IsDisposed)
                {
                    lblUpdateLatestDate.Text = "Không kiểm tra được";
                    lblUpdateStatus.Text = "Không kiểm tra được phiên bản: " + ex.Message;
                }
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
                lblUpdateStatus.Text = "Đang tải file ISO...";

                try
                {
                    using (var response = await UpdateHttpClient.GetAsync(
                        DownloadUrl, HttpCompletionOption.ResponseHeadersRead))
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

                    progressUpdate.Style = ProgressBarStyle.Blocks;
                    progressUpdate.Value = 100;
                    lblUpdateStatus.Text = "Đã tải xong file cập nhật.";
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
    }
}
