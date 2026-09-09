using System;

namespace SealAddIn.Models
{
    /// <summary>
    /// 電子押印(スタンプ配置)の簡易操作ログ1件を表します。
    /// 法的効力のある電子署名ではなく、誰がいつどこに押印したかを記録する簡易ログです。
    /// </summary>
    public class StampLogRecord
    {
        public string Id { get; set; }
        public string UserName { get; set; }

        public string WorkbookName { get; set; }
        public string WorksheetName { get; set; }
        public string CellAddress { get; set; }
        public string ShapeName { get; set; }
        public string StampText { get; set; }
        public string StampColor { get; set; }
        public DateTime StampedAt { get; set; }
    }
}
