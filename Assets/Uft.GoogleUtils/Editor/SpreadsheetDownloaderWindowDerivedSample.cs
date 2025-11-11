using System;
using System.IO;
using UnityEditor;

namespace Uft.GoogleUtils
{
    public class SpreadsheetDownloaderWindowDerivedSample : SpreadsheetDownloaderWindowBase
    {
        const string TITLE = "Spreadsheet Downloader (DerivedSample)";

        [MenuItem("Tools/Uft.GoogleUtils/" + TITLE, priority = 21100, secondaryPriority = 20)]
        public static void Open2()
        {
            var window = GetWindow<SpreadsheetDownloaderWindowDerivedSample>(TITLE);
            window.sheetUrl = "https://docs.google.com/spreadsheets/d/1yr_XvnNwMYrADDH_v4caCumOFEfYsI5mBXB20HapJrQ/edit?gid=0#gid=0";
            window.downloadDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            window.outputDirectory = "Assets/";
            window.outputFileName = "SpreadsheetDownloaderWindowDerivedSample.csv";
            window.overwritesExisting = true;
            window.browserOptions = new[] { "Default browser", "chrome.exe", "msedge.exe" };
            window.selectedBrowserIndex = 0;
            window.timeout_sec = 15;
            window.status = "";
        }
    }
}
