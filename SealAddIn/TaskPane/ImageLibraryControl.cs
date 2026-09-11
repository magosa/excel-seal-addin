using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using SealAddIn.Models;
using SealAddIn.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace SealAddIn.TaskPane
{
    /// <summary>画像ライブラリ(個人/共有)のアップロード・一覧表示・シートへの配置・削除を行うタブ。</summary>
    public class ImageLibraryControl : UserControl
    {
        private string currentScope = "personal";
        private TextBox displayNameTextBox;
        private Button scopePersonalButton;
        private Button scopeSharedButton;
        private FlowLayoutPanel grid;
        private Label statusLabel;

        public ImageLibraryControl()
        {
            BuildUi();
            SetScope("personal");
        }

        private void BuildUi()
        {
            Dock = DockStyle.Fill;
            // Officeのグレー系/ダークテーマ配下でも文字が読めるよう、テーマに依存しない配色を明示する。
            BackColor = Color.White;
            ForeColor = Color.Black;

            var uploadRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(6, 6, 6, 0) };
            uploadRow.Controls.Add(new Label { Text = "表示名(省略可):", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 4, 0) });
            displayNameTextBox = new TextBox { Width = 160 };
            uploadRow.Controls.Add(displayNameTextBox);
            var uploadButton = new Button { Text = "画像を選択してアップロード", AutoSize = true };
            uploadButton.Click += (s, e) => UploadImage();
            uploadRow.Controls.Add(uploadButton);

            var scopeRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(6, 4, 6, 4) };
            scopePersonalButton = new Button { Text = "個人ライブラリ", AutoSize = true };
            scopePersonalButton.Click += (s, e) => SetScope("personal");
            scopeSharedButton = new Button { Text = "共有ライブラリ", AutoSize = true };
            scopeSharedButton.Click += (s, e) => SetScope("shared");
            var refreshButton = new Button { Text = "再読み込み", AutoSize = true };
            refreshButton.Click += (s, e) => RefreshList();
            scopeRow.Controls.Add(scopePersonalButton);
            scopeRow.Controls.Add(scopeSharedButton);
            scopeRow.Controls.Add(refreshButton);

            statusLabel = new Label { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(6, 0, 6, 4), ForeColor = Color.DimGray };

            grid = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(6) };

            Controls.Add(grid);
            Controls.Add(statusLabel);
            Controls.Add(scopeRow);
            Controls.Add(uploadRow);
        }

        private void SetScope(string scope)
        {
            currentScope = scope;
            scopePersonalButton.FlatStyle = scope == "personal" ? FlatStyle.Flat : FlatStyle.Standard;
            scopeSharedButton.FlatStyle = scope == "shared" ? FlatStyle.Flat : FlatStyle.Standard;
            RefreshList();
        }

        public void RefreshList()
        {
            try
            {
                var records = ImageLibraryStorage.ListImages(currentScope);
                RenderList(records);
                statusLabel.Text = string.Empty;
            }
            catch (Exception ex)
            {
                grid.Controls.Clear();
                statusLabel.Text = ex.Message;
            }
        }

        private void RenderList(System.Collections.Generic.List<ImageRecord> records)
        {
            grid.Controls.Clear();

            if (records == null || records.Count == 0)
            {
                statusLabel.Text = "画像がまだ登録されていません。";
                return;
            }

            foreach (var record in records)
            {
                grid.Controls.Add(BuildItem(record));
            }
        }

        private Control BuildItem(ImageRecord record)
        {
            var panel = new Panel { Width = 140, Height = 190, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(4) };

            var pictureBox = new PictureBox { Width = 128, Height = 100, Top = 4, Left = 4, SizeMode = PictureBoxSizeMode.Zoom, BorderStyle = BorderStyle.FixedSingle };
            try
            {
                var path = ImageLibraryStorage.GetImagePath(record.Id, record.Scope);
                if (path != null)
                {
                    using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                    {
                        pictureBox.Image = Image.FromStream(fs);
                    }
                }
            }
            catch
            {
                // サムネイル読み込みに失敗しても一覧表示は継続します。
            }

            var nameLabel = new Label { Text = record.DisplayName, Top = 108, Left = 4, Width = 128, Height = 32, AutoEllipsis = true };

            var placeButton = new Button { Text = "配置", Top = 142, Left = 4, Width = 60, Height = 24 };
            placeButton.Click += (s, e) => PlaceImageOnSheet(record);

            var deleteButton = new Button { Text = "削除", Top = 142, Left = 68, Width = 60, Height = 24 };
            deleteButton.Click += (s, e) => DeleteImage(record);

            panel.Controls.Add(pictureBox);
            panel.Controls.Add(nameLabel);
            panel.Controls.Add(placeButton);
            panel.Controls.Add(deleteButton);
            return panel;
        }

        private void UploadImage()
        {
            using (var dialog = new OpenFileDialog { Filter = "画像ファイル (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg" })
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    var data = File.ReadAllBytes(dialog.FileName);
                    var ext = Path.GetExtension(dialog.FileName).ToLowerInvariant();
                    var contentType = ext == ".png" ? "image/png" : "image/jpeg";
                    var displayName = string.IsNullOrWhiteSpace(displayNameTextBox.Text)
                        ? Path.GetFileNameWithoutExtension(dialog.FileName)
                        : displayNameTextBox.Text.Trim();

                    ImageLibraryStorage.SaveImage(data, contentType, displayName, currentScope);
                    displayNameTextBox.Text = string.Empty;
                    RefreshList();
                    MessageBox.Show(this, "画像をライブラリに保存しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void PlaceImageOnSheet(ImageRecord record)
        {
            try
            {
                var path = ImageLibraryStorage.GetImagePath(record.Id, record.Scope);
                if (path == null)
                {
                    MessageBox.Show(this, "画像ファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var app = Globals.ThisAddIn.Application;
                var sheet = (Excel.Worksheet)app.ActiveSheet;
                var selection = (Excel.Range)app.Selection;
                float left = selection != null ? (float)(double)selection.Left : 0f;
                float top = selection != null ? (float)(double)selection.Top : 0f;

                var shape = sheet.Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, left, top, -1, -1);
                shape.Name = "IMG_" + DateTime.Now.Ticks.ToString("x");

                MessageBox.Show(this, "画像をシートに配置しました。「サイズ・位置編集」タブから調整できます。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteImage(ImageRecord record)
        {
            var confirm = MessageBox.Show(this, record.DisplayName + " を削除しますか?", "削除の確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes)
            {
                return;
            }

            try
            {
                ImageLibraryStorage.DeleteImage(record.Id, record.Scope, IdentityService.GetOwnerId());
                RefreshList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
