using System.Windows.Forms;

namespace SealAddIn.Options
{
    /// <summary>共有ライブラリ・押印ログの保存先フォルダーや表示名を設定するオプション画面。</summary>
    public class OptionsForm : Form
    {
        private readonly TextBox sharedLibraryTextBox;
        private readonly TextBox stampLogTextBox;
        private readonly TextBox displayNameTextBox;

        public OptionsForm()
        {
            Text = "オプション設定";
            Width = 520;
            Height = 260;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;

            var buttonPanel = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new System.Windows.Forms.Padding(8) };
            var cancelButton = new Button { Text = "キャンセル", DialogResult = DialogResult.Cancel, AutoSize = true };
            var okButton = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
            okButton.Click += (s, e) => Save();
            buttonPanel.Controls.Add(cancelButton);
            buttonPanel.Controls.Add(okButton);
            AcceptButton = okButton;
            CancelButton = cancelButton;

            var content = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 3, Padding = new System.Windows.Forms.Padding(12) };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70));

            content.Controls.Add(new Label { Text = "共有ライブラリ フォルダー:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            sharedLibraryTextBox = new TextBox { Dock = DockStyle.Fill, Text = Properties.Settings.Default.SharedLibraryPath };
            content.Controls.Add(sharedLibraryTextBox, 1, 0);
            var browseShared = new Button { Text = "参照...", Dock = DockStyle.Fill };
            browseShared.Click += (s, e) => BrowseFolder(sharedLibraryTextBox);
            content.Controls.Add(browseShared, 2, 0);

            content.Controls.Add(new Label { Text = "押印ログ フォルダー:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
            stampLogTextBox = new TextBox { Dock = DockStyle.Fill, Text = Properties.Settings.Default.StampLogFolderPath };
            content.Controls.Add(stampLogTextBox, 1, 1);
            var browseLog = new Button { Text = "参照...", Dock = DockStyle.Fill };
            browseLog.Click += (s, e) => BrowseFolder(stampLogTextBox);
            content.Controls.Add(browseLog, 2, 1);

            content.Controls.Add(new Label { Text = "表示名(任意):", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
            displayNameTextBox = new TextBox { Dock = DockStyle.Fill, Text = Properties.Settings.Default.DisplayNameOverride };
            content.Controls.Add(displayNameTextBox, 1, 2);

            Controls.Add(content);
            Controls.Add(buttonPanel);
        }

        private static void BrowseFolder(TextBox target)
        {
            using (var dialog = new FolderBrowserDialog { SelectedPath = target.Text })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    target.Text = dialog.SelectedPath;
                }
            }
        }

        private void Save()
        {
            Properties.Settings.Default.SharedLibraryPath = sharedLibraryTextBox.Text.Trim();
            Properties.Settings.Default.StampLogFolderPath = stampLogTextBox.Text.Trim();
            Properties.Settings.Default.DisplayNameOverride = displayNameTextBox.Text.Trim();
            Properties.Settings.Default.Save();
        }
    }
}
