using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using SealAddIn.Models;

namespace SealAddIn.Services
{
    /// <summary>
    /// 電子押印(スタンプ配置)の簡易操作ログをファイルに追記・取得します。
    /// 複数PCの複数プロセスが同じ共有フォルダーへ同時に書き込む構成になるため、
    /// ネットワーク越しの同時書き込み競合を避けるためにユーザー単位でログファイルを分割します。
    /// (例: \\server\share\StampLogs\202609_yamada.taro.jsonl)
    /// </summary>
    public static class StampLogStorage
    {
        private static readonly object WriteLock = new object();

        private static string RootPath => Properties.Settings.Default.StampLogFolderPath;

        private static string GetLogFilePath(DateTime dateUtc, string userName)
        {
            if (string.IsNullOrWhiteSpace(RootPath))
            {
                throw new InvalidOperationException("押印ログの保存先フォルダーが設定されていません。オプションで設定してください。");
            }
            Directory.CreateDirectory(RootPath);
            var fileName = dateUtc.ToString("yyyyMM") + "_" + PathUtil.SanitizeFileNameSegment(userName) + ".jsonl";
            return Path.Combine(RootPath, fileName);
        }

        public static StampLogRecord AppendLog(StampLogRecord record)
        {
            if (record == null || string.IsNullOrWhiteSpace(record.UserName))
            {
                throw new ArgumentException("押印者情報が取得できませんでした。");
            }

            record.Id = Guid.NewGuid().ToString("N");
            record.StampedAt = DateTime.UtcNow;

            var serializer = new JavaScriptSerializer();
            var line = serializer.Serialize(record);

            lock (WriteLock)
            {
                File.AppendAllText(GetLogFilePath(record.StampedAt, record.UserName), line + Environment.NewLine);
            }

            return record;
        }

        /// <summary>
        /// 直近2か月分のログファイル(全ユーザー分)をマージし、新しい順に取得します。
        /// </summary>
        public static List<StampLogRecord> GetRecentLogs(string workbookName, string worksheetName, int take)
        {
            var result = new List<StampLogRecord>();
            if (string.IsNullOrWhiteSpace(RootPath) || !Directory.Exists(RootPath))
            {
                return result;
            }

            var serializer = new JavaScriptSerializer();
            var targetMonths = new[] { DateTime.UtcNow.ToString("yyyyMM"), DateTime.UtcNow.AddMonths(-1).ToString("yyyyMM") };

            var candidates = new List<StampLogRecord>();
            var files = Directory.GetFiles(RootPath, "*.jsonl")
                .Where(f => targetMonths.Any(m => Path.GetFileName(f).StartsWith(m, StringComparison.Ordinal)));

            foreach (var file in files)
            {
                foreach (var line in File.ReadAllLines(file))
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    StampLogRecord record;
                    try
                    {
                        record = serializer.Deserialize<StampLogRecord>(line);
                    }
                    catch
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(workbookName) && record.WorkbookName != workbookName)
                    {
                        continue;
                    }
                    if (!string.IsNullOrEmpty(worksheetName) && record.WorksheetName != worksheetName)
                    {
                        continue;
                    }

                    candidates.Add(record);
                }
            }

            return candidates.OrderByDescending(r => r.StampedAt).Take(take).ToList();
        }
    }
}
