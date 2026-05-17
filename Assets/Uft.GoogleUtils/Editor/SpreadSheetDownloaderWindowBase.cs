#nullable enable

using System;
using System.IO;
using UnityEditor;
using UnityEngine;
using Debug = UnityEngine.Debug;

namespace Uft.GoogleUtils
{
    /// <summary>継承して各フィールドの値の初期値を入れたバージョンを用意すると便利です。</summary>
    public class SpreadsheetDownloaderWindowBase : EditorWindow
    {
        const string TITLE = "Sheet Downloader";

        public static void Open<T>() where T : EditorWindow
        {
            var all = Resources.FindObjectsOfTypeAll<T>();
            T? window = null;
            for (int i = 0; i < all.Length; i++)
            {
                if (all[i].GetType() == typeof(T))
                {
                    window = all[i];
                    break;
                }
            }
            if (window == null)
            {
                window = CreateInstance<T>();
            }
            window.Show();
            window.Focus();
        }

        [MenuItem("Tools/Uft.GoogleUtils/" + TITLE, priority = 21062000 + 715, secondaryPriority = 10)]
        public static void Open() => Open<SpreadsheetDownloaderWindowBase>();

        protected virtual void OnEnable()
        {
            this.titleContent.text = TITLE;
            this.sheetUrl = "";
            this.downloadDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            this.outputDirectory = "Assets/";
            this.outputFileName = "sample.csv";
            this.overwritesExisting = false;
            this.browserOptions = new[] { "Default browser", "chrome.exe", "msedge.exe" };
            this.selectedBrowserIndex = 0;
            this.timeout_sec = 15;
            this.status = "";
        }

        protected string? sheetUrl;
        protected string? downloadDirectory;
        protected string? outputDirectory;
        protected string? outputFileName;
        protected bool overwritesExisting;
        protected string[]? browserOptions;
        protected int selectedBrowserIndex;
        protected int timeout_sec;
        protected string? status;

        protected Color backGroundColor;

        bool _isRunning = false;

        protected virtual void OnGUI()
        {
            this.minSize = new Vector2(600, 280);
            this.maxSize = new Vector2(1800, 280);
            using (new EditorGUI.DisabledGroupScope(this._isRunning))
            {
                EditorGUI.DrawRect(new Rect(0, 0, this.position.width, this.position.height), backGroundColor);

                GUILayout.Label("ブラウザ経由で Spreadsheet を CSV ダウンロードします", EditorStyles.boldLabel);

                GUILayout.Space(10);

                this.sheetUrl = EditorGUILayout.TextField("シート URL", this.sheetUrl);
                this.downloadDirectory = EditorGUILayout.TextField("ダウンロードフォルダ", this.downloadDirectory);
                this.outputDirectory = EditorGUILayout.TextField("出力先フォルダ", this.outputDirectory);
                this.outputFileName = EditorGUILayout.TextField("出力ファイル名", this.outputFileName);
                this.overwritesExisting = EditorGUILayout.Toggle("上書き許可", this.overwritesExisting);
                this.selectedBrowserIndex = EditorGUILayout.Popup("ブラウザ", this.selectedBrowserIndex, this.browserOptions);
                this.timeout_sec = EditorGUILayout.IntSlider("タイムアウト(秒)", this.timeout_sec, 15, 30);

                GUILayout.Space(10);

                if (GUILayout.Button("Download CSV 📥"))
                {
                    async void taskVoid()
                    {
                        if (this._isRunning) return;
                        SpreadsheetDownloader? downloader = null;
                        try
                        {
                            this._isRunning = true;
                            this.status = "ダウンロード中";
                            Debug.Log(this.status);
                            this.Repaint();
                            downloader = new SpreadsheetDownloader(
                                this.downloadDirectory!,
                                this.outputDirectory!,
                                this.outputFileName!,
                                this.overwritesExisting!,
                                this.selectedBrowserIndex == 0 ? null : this.browserOptions![this.selectedBrowserIndex],
                                TimeSpan.FromSeconds(this.timeout_sec));
                            var csvUrl = SpreadsheetDownloader.GetCsvExportUrl(this.sheetUrl!);
                            var destPath = await downloader.DownloadCsvAsync(csvUrl);
                            this.status = $"✅ ダウンロード完了: {destPath}";
                            Debug.Log(this.status);
                            AssetDatabase.Refresh();
                        }
                        catch (Exception ex)
                        {
                            this.status = $"💥{ex.Message}";
                            Debug.LogError(this.status);
                        }
                        finally
                        {
                            this._isRunning = false;
                            this.Repaint();
                            downloader?.Dispose();
                        }
                    }
                    taskVoid();
                }
                GUILayout.Space(10);
                if (GUILayout.Button("Clear status")) this.status = "";
                GUILayout.Label(this.status, EditorStyles.boldLabel);
            }
        }
    }
}
