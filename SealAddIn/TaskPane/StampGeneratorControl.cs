using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using SealAddIn.Models;
using SealAddIn.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace SealAddIn.TaskPane
{
    /// <summary>印影(スタンプ)を生成し、選択セルへ配置する(電子押印する)タブ。</summary>
    public class StampGeneratorControl : UserControl
    {
        private TextBox textInput;
        private ComboBox shapeComboBox;
        private Button colorButton;
        private NumericUpDown sizeInput;
        private PictureBox previewBox;
        private ComboBox saveScopeComboBox;
        private Color selectedColor = ColorTranslator.FromHtml("#cc0000");

        public StampGeneratorControl()
        {
            BuildUi();
            UpdatePreview();
        }

        private void BuildUi()
        {
            Dock = DockStyle.Fill;
            var layout = new FlowLayoutPanel { Dock = DockStyle.Top, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Padding = new Padding(6) };

            var row1 = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            row1.Controls.Add(new Label { Text = "氏名/部署名:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) });
            textInput = new TextBox { Width = 120, MaxLength = 6 };
            textInput.TextChanged += (s, e) => UpdatePreview();
            row1.Controls.Add(textInput);

            var row2 = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            row2.Controls.Add(new Label { Text = "形状:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) });
            shapeComboBox = new ComboBox { Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            shapeComboBox.Items.AddRange(new object[] { "丸印", "角印" });
            shapeComboBox.SelectedIndex = 0;
            shapeComboBox.SelectedIndexChanged += (s, e) => UpdatePreview();
            row2.Controls.Add(shapeComboBox);

            var row3 = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            row3.Controls.Add(new Label { Text = "色:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) });
            colorButton = new Button { Width = 60, Height = 24, BackColor = selectedColor };
            colorButton.Click += (s, e) => PickColor();
            row3.Controls.Add(colorButton);

            var row4 = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            row4.Controls.Add(new Label { Text = "サイズ(mm):", AutoSize = true, Margin = new Padding(0, 6, 4, 0) });
            sizeInput = new NumericUpDown { Width = 70, Minimum = 5, Maximum = 60, Value = 18 };
            sizeInput.ValueChanged += (s, e) => UpdatePreview();
            row4.Controls.Add(sizeInput);

            var placeButton = new Button { Text = "選択セルに押印", AutoSize = true, Margin = new Padding(0, 6, 0, 6) };
            placeButton.Click += (s, e) => PlaceStamp();

            var row5 = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            saveScopeComboBox = new ComboBox { Width = 140, DropDownStyle = ComboBoxStyle.DropDownList };
            saveScopeComboBox.Items.AddRange(new object[] { "個人ライブラリへ", "共有ライブラリへ" });
            saveScopeComboBox.SelectedIndex = 0;
            row5.Controls.Add(saveScopeComboBox);
            var saveButton = new Button { Text = "ライブラリに保存", AutoSize = true };
            saveButton.Click += (s, e) => SaveToLibrary();
            row5.Controls.Add(saveButton);

            previewBox = new PictureBox { Width = 200, Height = 200, SizeMode = PictureBoxSizeMode.Zoom, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 8, 0, 0) };

            layout.Controls.Add(row1);
            layout.Controls.Add(row2);
            layout.Controls.Add(row3);
            layout.Controls.Add(row4);
            layout.Controls.Add(placeButton);
            layout.Controls.Add(row5);
            layout.Controls.Add(previewBox);

            Controls.Add(layout);
        }

        private void PickColor()
        {
            using (var dialog = new ColorDialog { Color = selectedColor })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    selectedColor = dialog.Color;
                    colorButton.BackColor = selectedColor;
                    UpdatePreview();
                }
            }
        }

        private string ShapeKey => shapeComboBox.SelectedIndex == 1 ? "rect" : "circle";

        private string ColorHex => ColorTranslator.ToHtml(selectedColor);

        private void UpdatePreview()
        {
            try
            {
                var data = StampImageRenderer.Render(textInput.Text.Trim(), ShapeKey, ColorHex);
                using (var ms = new MemoryStream(data))
                {
                    previewBox.Image?.Dispose();
                    previewBox.Image = Image.FromStream(ms);
                }
            }
            catch
            {
                // プレビュー生成エラーは無視します(入力途中で発生しうるため)。
            }
        }

        private void PlaceStamp()
        {
            var text = textInput.Text.Trim();
            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show(this, "氏名(または部署名)を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string tempFile = null;
            try
            {
                var data = StampImageRenderer.Render(text, ShapeKey, ColorHex);
                tempFile = Path.Combine(Path.GetTempPath(), "SealAddIn_" + Guid.NewGuid().ToString("N") + ".png");
                File.WriteAllBytes(tempFile, data);

                var app = Globals.ThisAddIn.Application;
                var workbook = app.ActiveWorkbook;
                var sheet = (Excel.Worksheet)app.ActiveSheet;
                var selection = (Excel.Range)app.Selection;
                float left = selection != null ? (float)(double)selection.Left : 0f;
                float top = selection != null ? (float)(double)selection.Top : 0f;
                var cellAddress = selection != null
                    ? selection.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing)
                    : string.Empty;

                var sizePt = (float)UnitConversion.MmToPoints((double)sizeInput.Value);
                var shapeName = "STAMP_" + DateTime.Now.Ticks.ToString("x");

                var shape = sheet.Shapes.AddPicture(tempFile, Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, left, top, sizePt, sizePt);
                shape.Name = shapeName;

                try
                {
                    StampLogStorage.AppendLog(new StampLogRecord
                    {
                        UserName = IdentityService.GetDisplayName(),
                        WorkbookName = workbook != null ? workbook.Name : string.Empty,
                        WorksheetName = sheet.Name,
                        CellAddress = cellAddress,
                        ShapeName = shapeName,
                        StampText = text,
                        StampColor = ColorHex
                    });
                }
                catch (InvalidOperationException logEx)
                {
                    MessageBox.Show(this, "印影を配置しましたが、押印ログは記録されませんでした。\n" + logEx.Message,
                        "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                MessageBox.Show(this, "印影をシートに配置し、押印ログを記録しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (tempFile != null && File.Exists(tempFile))
                {
                    try { File.Delete(tempFile); } catch { /* 一時ファイル削除失敗は無視 */ }
                }
            }
        }

        private void SaveToLibrary()
        {
            var text = textInput.Text.Trim();
            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show(this, "氏名(または部署名)を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var data = StampImageRenderer.Render(text, ShapeKey, ColorHex);
                var scope = saveScopeComboBox.SelectedIndex == 1 ? "shared" : "personal";
                ImageLibraryStorage.SaveImage(data, "image/png", text + "の印影", scope);
                MessageBox.Show(this, "生成した印影を画像ライブラリに保存しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
