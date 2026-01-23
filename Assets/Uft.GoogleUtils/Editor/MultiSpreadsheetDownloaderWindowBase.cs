#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using UnityEditor;
using UnityEngine;
using Debug = UnityEngine.Debug;

namespace Uft.GoogleUtils
{
    [Serializable]
    public class TargetInfo
    {
        public string saveName = "sample.csv";
        public string sheetUrl = "";
        [NonSerialized] public string status = "";
    }

    /// <summary>継承して各フィールドの値の初期値を入れたバージョンを用意すると便利です。</summary>
    public class MultiSpreadsheetDownloaderWindowBase : EditorWindow
    {
        const string TITLE = "Multi Sheet Downloader";

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

        [MenuItem("Tools/Uft.GoogleUtils/" + TITLE, priority = 21100, secondaryPriority = 30)]
        public static void Open() => Open<MultiSpreadsheetDownloaderWindowBase>();

        protected virtual void OnEnable()
        {
            this.titleContent.text = TITLE;

            this.downloadDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            this.outputDirectory = "Assets/";

            this.overwritesExisting = false;
            this.browserOptions = new[] { "Default browser", "chrome.exe", "msedge.exe" };
            this.selectedBrowserIndex = 0;
            this.timeout_sec = 15;
            this.status = "";
        }

        protected string? downloadDirectory;
        protected string? outputDirectory;

        protected List<TargetInfo> targetInfoList = new();

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
            this.maxSize = new Vector2(1800, 1000);
            using (new EditorGUI.DisabledGroupScope(this._isRunning))
            {
                EditorGUI.DrawRect(new Rect(0, 0, this.position.width, this.position.height), backGroundColor);

                GUILayout.Label("ブラウザ経由で Spreadsheet を CSV ダウンロードします", EditorStyles.boldLabel);

                GUILayout.Space(10);

                this.downloadDirectory = EditorGUILayout.TextField("ダウンロードフォルダ", this.downloadDirectory);
                this.outputDirectory = EditorGUILayout.TextField("出力先フォルダ", this.outputDirectory);
                this.DrawTargetInfoList();
                this.overwritesExisting = EditorGUILayout.Toggle("上書き許可", this.overwritesExisting);
                this.selectedBrowserIndex = EditorGUILayout.Popup("ブラウザ", this.selectedBrowserIndex, this.browserOptions);
                this.timeout_sec = EditorGUILayout.IntSlider("タイムアウト(秒)", this.timeout_sec, 15, 30);

                GUILayout.Space(10);

                if (GUILayout.Button("Download All CSV 📥"))
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

                            int i = 0;
                            for (i = 0; i < this.targetInfoList.Count; i++)
                            {
                                var t = this.targetInfoList[i];
                                if (downloader != null)
                                {
                                    downloader.Dispose();
                                    downloader = null;
                                }
                                downloader = new SpreadsheetDownloader(
                                    this.downloadDirectory!,
                                    this.outputDirectory!,
                                    t.saveName!,
                                    this.overwritesExisting!,
                                    this.selectedBrowserIndex == 0 ? null : this.browserOptions![this.selectedBrowserIndex],
                                    TimeSpan.FromSeconds(this.timeout_sec));
                                var csvUrl = SpreadsheetDownloader.GetCsvExportUrl(t.sheetUrl);
                                var destPath = await downloader.DownloadCsvAsync(csvUrl);
                                this.status = $"{i}/{this.targetInfoList.Count} 完了";
                                Debug.Log(this.status);

                            }
                            this.status = $"✅ ダウンロード完了: {i}件";
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

        protected virtual void DrawTargetInfoList()
        {
            GUILayout.Label("URLと保存名");
            for (int i = 0; i < this.targetInfoList.Count; i++)
            {
                var t = this.targetInfoList[i];
                using (new EditorGUILayout.HorizontalScope())
                {
                    GUILayout.Label($"[{i}]", GUILayout.Width(20));
                    var prev = EditorGUIUtility.labelWidth;
                    EditorGUIUtility.labelWidth = 20;
                    t.sheetUrl = EditorGUILayout.TextField("url", t.sheetUrl);
                    GUILayout.Space(10);
                    EditorGUIUtility.labelWidth = 30;
                    t.saveName = EditorGUILayout.TextField("save", t.saveName, GUILayout.Width(120));
                    EditorGUIUtility.labelWidth = prev;
                    if (GUILayout.Button("Download", GUILayout.Width(80)))
                    {
                        this.DownLoadOne(t.saveName, t.sheetUrl);
                    }
                    if (GUILayout.Button("-", GUILayout.Width(30)))
                    {
                        this.targetInfoList.RemoveAt(i);
                        GUI.FocusControl(null);
                        break;
                    }
                }
            }
            if (GUILayout.Button("+", GUILayout.Width(30)))
            {
                this.targetInfoList.Add(new TargetInfo());
            }
        }

        void DownLoadOne(string outputFileName, string sheetUrl)
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
                        outputFileName,
                        this.overwritesExisting!,
                        this.selectedBrowserIndex == 0 ? null : this.browserOptions![this.selectedBrowserIndex],
                        TimeSpan.FromSeconds(this.timeout_sec));
                    var csvUrl = SpreadsheetDownloader.GetCsvExportUrl(sheetUrl);
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
    }
}
