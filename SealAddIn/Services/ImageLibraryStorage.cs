using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using SealAddIn.Models;

namespace SealAddIn.Services
{
    /// <summary>
    /// 画像ライブラリ(個人/共有)のファイルシステムへの保存・読み込みを担当します。
    /// 個人ライブラリは %APPDATA% 配下(OSのユーザープロファイルで既にユーザー分離済み)、
    /// 共有ライブラリはオプションで設定した共有フォルダー(UNCパス等)に保存します。
    /// </summary>
    public static class ImageLibraryStorage
    {
        private const long MaxFileSizeBytes = 5 * 1024 * 1024; // 5MB
        private static readonly string[] AllowedContentTypes = { "image/png", "image/jpeg" };

        private static string PersonalRootPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SealAddIn", "ImageLibrary");

        private static string SharedRootPath => Properties.Settings.Default.SharedLibraryPath;

        private static string GetScopeFolder(string scope)
        {
            if (scope == "shared")
            {
                if (string.IsNullOrWhiteSpace(SharedRootPath))
                {
                    throw new InvalidOperationException("共有ライブラリの保存先フォルダーが設定されていません。オプションで設定してください。");
                }
                return SharedRootPath;
            }
            if (scope == "personal")
            {
                return PersonalRootPath;
            }
            throw new ArgumentException("保存先(scope)は personal または shared を指定してください。");
        }

        public static ImageRecord SaveImage(byte[] data, string contentType, string displayName, string scope)
        {
            if (data == null || data.Length == 0)
            {
                throw new ArgumentException("画像データが空です。");
            }
            if (data.Length > MaxFileSizeBytes)
            {
                throw new ArgumentException("画像サイズが上限(5MB)を超えています。");
            }
            if (!AllowedContentTypes.Contains(contentType))
            {
                throw new ArgumentException("PNGまたはJPEG形式の画像のみアップロードできます。");
            }

            int width, height;
            using (var ms = new MemoryStream(data))
            using (var image = Image.FromStream(ms))
            {
                width = image.Width;
                height = image.Height;
            }

            var folder = GetScopeFolder(scope);
            Directory.CreateDirectory(folder);

            var id = Guid.NewGuid().ToString("N");
            var extension = contentType == "image/png" ? ".png" : ".jpg";
            var fileName = id + extension;

            File.WriteAllBytes(Path.Combine(folder, fileName), data);

            var record = new ImageRecord
            {
                Id = id,
                DisplayName = string.IsNullOrWhiteSpace(displayName) ? fileName : displayName,
                FileName = fileName,
                ContentType = contentType,
                Scope = scope,
                OwnerId = IdentityService.GetOwnerId(),
                Width = width,
                Height = height,
                SizeBytes = data.Length,
                CreatedAt = DateTime.UtcNow
            };

            var serializer = new JavaScriptSerializer();
            File.WriteAllText(Path.Combine(folder, id + ".json"), serializer.Serialize(record));

            return record;
        }

        public static List<ImageRecord> ListImages(string scope)
        {
            var folder = GetScopeFolder(scope);
            var result = new List<ImageRecord>();
            if (!Directory.Exists(folder))
            {
                return result;
            }

            var serializer = new JavaScriptSerializer();
            foreach (var jsonPath in Directory.GetFiles(folder, "*.json"))
            {
                try
                {
                    var record = serializer.Deserialize<ImageRecord>(File.ReadAllText(jsonPath));
                    result.Add(record);
                }
                catch
                {
                    // 破損したメタデータはスキップします。
                }
            }

            return result.OrderByDescending(r => r.CreatedAt).ToList();
        }

        private static ImageRecord FindRecord(string id, string scope)
        {
            var folder = GetScopeFolder(scope);
            var jsonPath = Path.Combine(folder, id + ".json");
            if (!File.Exists(jsonPath))
            {
                return null;
            }
            var serializer = new JavaScriptSerializer();
            return serializer.Deserialize<ImageRecord>(File.ReadAllText(jsonPath));
        }

        /// <summary>画像配置のために、保存済み画像のフルパスを取得します。</summary>
        public static string GetImagePath(string id, string scope)
        {
            var record = FindRecord(id, scope);
            if (record == null)
            {
                return null;
            }
            var filePath = Path.Combine(GetScopeFolder(scope), record.FileName);
            return File.Exists(filePath) ? filePath : null;
        }

        /// <summary>
        /// 画像を削除します。共有ライブラリの画像であっても、アップロードした本人のみ削除できます。
        /// </summary>
        public static bool DeleteImage(string id, string scope, string requesterId)
        {
            var record = FindRecord(id, scope);
            if (record == null)
            {
                return false;
            }
            if (!string.Equals(record.OwnerId, requesterId, StringComparison.OrdinalIgnoreCase))
            {
                throw new UnauthorizedAccessException("アップロードした本人のみ削除できます。");
            }

            var folder = GetScopeFolder(scope);
            var filePath = Path.Combine(folder, record.FileName);
            var jsonPath = Path.Combine(folder, id + ".json");
            if (File.Exists(filePath)) File.Delete(filePath);
            if (File.Exists(jsonPath)) File.Delete(jsonPath);
            return true;
        }
    }
}
