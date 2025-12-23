#nullable enable

using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Uft.GoogleUtils
{
    public class SpreadsheetDownloader : IDisposable
    {
        // https://docs.google.com/spreadsheets/d/<ID>/edit?gid=0#gid=<GID>
        // ↓
        // https://docs.google.com/spreadsheets/d/<ID>/export?format=csv&gid=<GID>
        public static string GetCsvExportUrl(string sheetUrl)
        {
            if (!sheetUrl.Contains("/d/") ||
                !sheetUrl.Contains("gid="))
            {
                throw new ArgumentException($"Invalid spreadsheets url: {sheetUrl}", nameof(sheetUrl));
            }
            var parts = sheetUrl.Split(new string[] { "/d/" }, StringSplitOptions.None);
            var id = parts[1].Split('/')[0];
            var gidIndex = parts[1].IndexOf("gid=") + 4;
            var gid = parts[1][gidIndex..].Split(new char[] { '&', '#', '?' })[0];
            return $"{parts[0]}/d/{id}/export?format=csv&gid={gid}";
        }

        string _downloadDirectory;
        public string DownloadDirectory
        {
            get => this._downloadDirectory;
            set
            {
                if (!Directory.Exists(value))
                {
                    throw new ArgumentException($"Invalid path: {value}", nameof(value));
                }
                this._downloadDirectory = value;
            }
        }
        public string OutputDirectory { get; private set; }
        public string OutputFileName { get; private set; }
        public string OutputFilePath => Path.Combine(this.OutputDirectory, this.OutputFileName);
        public bool OverwritesExisting { get; private set; }
        public string? LaunchCommand { get; private set; }
        public TimeSpan Timeout { get; set; }

        readonly CancellationTokenSource clientLifetimeTokenSource = new();

        public SpreadsheetDownloader(string downloadDirectory, string outputDirectory, string outputFileName, bool overwritesExisting, string? launchCommand, TimeSpan timeout)
        {
            // HACK: Unity6 + C#9.0 の nullable解析が上手く行かないので、プロパティと同じ記述を書く
            // this.DownloadDirectory = downloadDirectory;
            if (!Directory.Exists(downloadDirectory))
            {
                throw new ArgumentException($"Invalid path: {downloadDirectory}", nameof(downloadDirectory));
            }
            this._downloadDirectory = downloadDirectory;

            this.OutputDirectory = outputDirectory;
            this.OutputFileName = outputFileName;
            this.OverwritesExisting = overwritesExisting;
            this.LaunchCommand = launchCommand;
            this.Timeout = timeout;
        }

        public async Task<string> DownloadCsvAsync(string csvExportUrl, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(csvExportUrl)) throw new ArgumentException($"{nameof(csvExportUrl)} is required.", nameof(csvExportUrl));

            if (string.IsNullOrWhiteSpace(this.LaunchCommand))
            {
                Process.Start(new ProcessStartInfo()
                {
                    FileName = csvExportUrl,
                    UseShellExecute = true, // NOTE: App Paths (registry。ここで言うShellとはWindowsの普段のマウス操作等のGUI操作相当のことを指す。) 経由でexeを見つけることを許可
                });
            }
            else
            {
                Process.Start(new ProcessStartInfo()
                {
                    FileName = this.LaunchCommand,
                    Arguments = $"\"{csvExportUrl}\"",
                    UseShellExecute = true, // NOTE: App Paths (registry。ここで言うShellとはWindowsの普段のマウス操作等のGUI操作相当のことを指す。) 経由でexeを見つけることを許可
                });
            }
            var latestFilePath = await this.WaitForDownloadCompletionAsync(cancellationToken);

            if (!Directory.Exists(this.OutputDirectory))
            {
                Directory.CreateDirectory(this.OutputDirectory);
            }
            if (this.OverwritesExisting && File.Exists(this.OutputFilePath)) File.Delete(this.OutputFilePath);
            File.Move(latestFilePath, this.OutputFilePath);
            return this.OutputFilePath;
        }

        async Task<string> WaitForDownloadCompletionAsync(CancellationToken cancellationToken)
        {
            var downloadDirectory = this.DownloadDirectory;

            // NOTE: .NET と ファイル更新日時の誤差を吸収
            var startTime_utc = DateTime.UtcNow.AddSeconds(-1);
            var timeout = this.Timeout + TimeSpan.FromSeconds(+1);
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(this.clientLifetimeTokenSource.Token, cancellationToken);
            cts.CancelAfter(timeout);

            // NOTE: おまけ
            if (startTime_utc.ToString("yyyyMMdd") == "20380119") throw new Exception("今日は2038年問題(2038-01-19T03:14:07Z)です。有給取得したらいかがですか？");

            try
            {
                while (true)
                {
                    await Task.Delay(1000, cts.Token);
                    var csvFiles = Directory
                        .GetFiles(downloadDirectory, "*.csv")
                        .Where(csv => !File.Exists(csv + ".crdownload") && File.GetLastWriteTimeUtc(csv) > startTime_utc)
                        .ToArray();
                    if (0 < csvFiles.Length)
                    {
                        await Task.Delay(1000, cts.Token); // NOTE: 追加で１秒待つ
                        return csvFiles
                            .OrderByDescending(csv => File.GetLastWriteTimeUtc(csv).Ticks)
                            .First();
                    }
                }
            }
            catch (OperationCanceledException ex) when (ex.CancellationToken == cts.Token)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    throw new OperationCanceledException(ex.Message, ex, cancellationToken);
                }
                else if (this.clientLifetimeTokenSource.IsCancellationRequested)
                {
                    throw new OperationCanceledException("Client is disposed.", ex, this.clientLifetimeTokenSource.Token);
                }
                else
                {
                    throw new TimeoutException($"The request was canceled due to the configured Timeout of {this.Timeout.TotalSeconds} seconds elapsing.", ex);
                }
            }
        }

        public void Dispose()
        {
            this.clientLifetimeTokenSource.Cancel();
            this.clientLifetimeTokenSource.Dispose();
        }
    }
}
