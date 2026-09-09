using System;

namespace SealAddIn.Models
{
    /// <summary>
    /// 画像ライブラリに保存された1件の画像のメタデータを表します。
    /// </summary>
    public class ImageRecord
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string FileName { get; set; }
        public string ContentType { get; set; }

        /// <summary>"personal" または "shared"。</summary>
        public string Scope { get; set; }

        public string OwnerId { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public long SizeBytes { get; set; }
        public DateTime CreatedAt { get; set; }
    }
}
