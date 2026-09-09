using System;
using System.Windows.Forms;
using SealAddIn.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace SealAddIn.TaskPane
{
    /// <summary>押印履歴(簡易ログ)を表示するタブ。</summary>
    public class StampHistoryControl : UserControl
    {
        private ListView listView;
        private Label statusLabel;

        public StampHistoryControl()
        {
            BuildUi();
        }

        private void BuildUi()
        {
            Dock = DockStyle.Fill;

            var topRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(6) };
            var refreshButton = new Button { Text = "再読み込み", AutoSize = true };
            refreshButton.Click += (s, e) => RefreshList();
            topRow.Controls.Add(refreshButton);

            statusLabel = new Label { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(6, 0, 6, 4), ForeColor = System.Drawing.Color.DimGray };

            listView = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true };
            listView.Columns.Add("押印者", 100);
            listView.Columns.Add("日時", 130);
            listView.Columns.Add("セル", 60);
            listView.Columns.Add("印影テキスト", 100);

            Controls.Add(listView);
            Controls.Add(statusLabel);
            Controls.Add(topRow);
        }

        public void RefreshList()
        {
            listView.Items.Clear();
            statusLabel.Text = string.Empty;

            try
            {
                var app = Globals.ThisAddIn.Application;
                var workbook = app.ActiveWorkbook;
                var sheet = (Excel.Worksheet)app.ActiveSheet;
                var workbookName = workbook != null ? workbook.Name : string.Empty;
                var worksheetName = sheet != null ? sheet.Name : string.Empty;

                var records = StampLogStorage.GetRecentLogs(workbookName, worksheetName, 20);
                if (records.Count == 0)
                {
                    statusLabel.Text = "このシートにはまだ押印履歴がありません。";
                    return;
                }

                foreach (var record in records)
                {
                    var item = new ListViewItem(record.UserName ?? "不明なユーザー");
                    item.SubItems.Add(record.StampedAt.ToLocalTime().ToString("yyyy/MM/dd HH:mm"));
                    item.SubItems.Add(record.CellAddress ?? string.Empty);
                    item.SubItems.Add(record.StampText ?? string.Empty);
                    listView.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                statusLabel.Text = ex.Message;
            }
        }
    }
}
