using System;
using System.IO;
using UnityEditor;

namespace Uft.GoogleUtils
{
    public class SpreadsheetDownloaderWindowDerivedSample : SpreadsheetDownloaderWindowBase
    {
        const string TITLE = "Sheet Downloader (DerivedSample)";

        [MenuItem("Tools/Uft.GoogleUtils/" + TITLE, priority = 21100, secondaryPriority = 20)]
        public static void Open2() => Open<SpreadsheetDownloaderWindowDerivedSample>();

        protected override void OnEnable()
        {
            this.titleContent.text = TITLE;
            this.sheetUrl = "https://docs.google.com/spreadsheets/d/1yr_XvnNwMYrADDH_v4caCumOFEfYsI5mBXB20HapJrQ/edit?gid=0#gid=0";
            this.downloadDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            this.outputDirectory = "Assets/";
            this.outputFileName = "SpreadsheetDownloaderWindowDerivedSample.csv";
            this.overwritesExisting = true;
            this.browserOptions = new[] { "Default browser", "chrome.exe", "msedge.exe" };
            this.selectedBrowserIndex = 0;
            this.timeout_sec = 15;
            this.status = "";
        }
    }
}
