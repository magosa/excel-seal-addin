using System;
using System.Collections.Generic;
using System.Windows.Forms;
using SealAddIn.Services;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace SealAddIn.TaskPane
{
    /// <summary>アクティブシートに配置済みの画像(図形)のサイズ・位置を編集するタブ。</summary>
    public class ShapeEditorControl : UserControl
    {
        private class ShapeInfo
        {
            public double Left, Top, Width, Height;
        }

        private readonly Dictionary<string, ShapeInfo> shapesByName = new Dictionary<string, ShapeInfo>();

        private ComboBox shapeListComboBox;
        private ComboBox unitComboBox;
        private TextBox widthTextBox, heightTextBox, leftTextBox, topTextBox;
        private CheckBox lockAspectCheckBox;

        public ShapeEditorControl()
        {
            BuildUi();
            RefreshShapeList();
        }

        private static FlowLayoutPanel RowPanel()
        {
            return new FlowLayoutPanel { AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.LeftToRight };
        }

        private void BuildUi()
        {
            Dock = DockStyle.Fill;
            var layout = new FlowLayoutPanel { Dock = DockStyle.Top, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Padding = new System.Windows.Forms.Padding(6) };

            var row1 = RowPanel();
            row1.Controls.Add(new Label { Text = "シート内の画像:", AutoSize = true, Margin = new System.Windows.Forms.Padding(0, 6, 4, 0) });
            shapeListComboBox = new ComboBox { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
            shapeListComboBox.SelectedIndexChanged += (s, e) => OnSelectShape();
            row1.Controls.Add(shapeListComboBox);
            var refreshBtn = new Button { Text = "一覧を更新", AutoSize = true };
            refreshBtn.Click += (s, e) => RefreshShapeList();
            row1.Controls.Add(refreshBtn);

            var row2 = RowPanel();
            row2.Controls.Add(new Label { Text = "単位:", AutoSize = true, Margin = new System.Windows.Forms.Padding(0, 6, 4, 0) });
            unitComboBox = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            unitComboBox.Items.AddRange(new object[] { "mm", "pt", "px" });
            unitComboBox.SelectedIndex = 0;
            unitComboBox.SelectedIndexChanged += (s, e) => OnSelectShape();
            row2.Controls.Add(unitComboBox);

            var row3 = RowPanel();
            row3.Controls.Add(new Label { Text = "幅:", AutoSize = true, Margin = new System.Windows.Forms.Padding(0, 6, 4, 0) });
            widthTextBox = new TextBox { Width = 70 };
            row3.Controls.Add(widthTextBox);
            row3.Controls.Add(new Label { Text = "高さ:", AutoSize = true, Margin = new System.Windows.Forms.Padding(8, 6, 4, 0) });
            heightTextBox = new TextBox { Width = 70 };
            row3.Controls.Add(heightTextBox);

            var row4 = RowPanel();
            lockAspectCheckBox = new CheckBox { Text = "縦横比を固定する", AutoSize = true };
            lockAspectCheckBox.CheckedChanged += (s, e) => heightTextBox.Enabled = !lockAspectCheckBox.Checked;
            row4.Controls.Add(lockAspectCheckBox);

            var row5 = RowPanel();
            row5.Controls.Add(new Label { Text = "左位置:", AutoSize = true, Margin = new System.Windows.Forms.Padding(0, 6, 4, 0) });
            leftTextBox = new TextBox { Width = 70 };
            row5.Controls.Add(leftTextBox);
            row5.Controls.Add(new Label { Text = "上位置:", AutoSize = true, Margin = new System.Windows.Forms.Padding(8, 6, 4, 0) });
            topTextBox = new TextBox { Width = 70 };
            row5.Controls.Add(topTextBox);

            var applyButton = new Button { Text = "適用", AutoSize = true, Margin = new System.Windows.Forms.Padding(0, 6, 0, 6) };
            applyButton.Click += (s, e) => ApplyChanges();

            var presetRow = RowPanel();
            presetRow.Controls.Add(new Label { Text = "よく使う印鑑サイズ:", AutoSize = true, Margin = new System.Windows.Forms.Padding(0, 6, 4, 0) });
            foreach (var mm in new[] { 15, 18, 21 })
            {
                var mmLocal = mm;
                var presetButton = new Button { Text = mm + "mm角", AutoSize = true };
                presetButton.Click += (s, e) => ApplyPreset(mmLocal);
                presetRow.Controls.Add(presetButton);
            }

            layout.Controls.Add(row1);
            layout.Controls.Add(row2);
            layout.Controls.Add(row3);
            layout.Controls.Add(row4);
            layout.Controls.Add(row5);
            layout.Controls.Add(applyButton);
            layout.Controls.Add(presetRow);

            Controls.Add(layout);
        }

        public void RefreshShapeList()
        {
            shapesByName.Clear();
            shapeListComboBox.Items.Clear();

            try
            {
                var app = Globals.ThisAddIn.Application;
                var sheet = (Excel.Worksheet)app.ActiveSheet;
                foreach (Excel.Shape shape in sheet.Shapes)
                {
                    shapesByName[shape.Name] = new ShapeInfo { Left = shape.Left, Top = shape.Top, Width = shape.Width, Height = shape.Height };
                    shapeListComboBox.Items.Add(shape.Name);
                }

                if (shapeListComboBox.Items.Count > 0)
                {
                    shapeListComboBox.SelectedIndex = 0;
                }
                else
                {
                    ClearInputs();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearInputs()
        {
            widthTextBox.Text = heightTextBox.Text = leftTextBox.Text = topTextBox.Text = string.Empty;
        }

        private string SelectedUnit => unitComboBox.SelectedItem as string ?? "mm";

        private void OnSelectShape()
        {
            var name = shapeListComboBox.SelectedItem as string;
            ShapeInfo info;
            if (name == null || !shapesByName.TryGetValue(name, out info))
            {
                ClearInputs();
                return;
            }

            var unit = SelectedUnit;
            widthTextBox.Text = Round2(UnitConversion.ConvertFromPoints(info.Width, unit)).ToString();
            heightTextBox.Text = Round2(UnitConversion.ConvertFromPoints(info.Height, unit)).ToString();
            leftTextBox.Text = Round2(UnitConversion.ConvertFromPoints(info.Left, unit)).ToString();
            topTextBox.Text = Round2(UnitConversion.ConvertFromPoints(info.Top, unit)).ToString();
        }

        private static double Round2(double value) => Math.Round(value, 2);

        private void ApplyChanges()
        {
            var name = shapeListComboBox.SelectedItem as string;
            if (name == null)
            {
                MessageBox.Show(this, "編集する画像を選択してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            double width, left, top;
            if (!double.TryParse(widthTextBox.Text, out width) || !double.TryParse(leftTextBox.Text, out left) || !double.TryParse(topTextBox.Text, out top))
            {
                MessageBox.Show(this, "数値を正しく入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            double height;
            var hasHeight = double.TryParse(heightTextBox.Text, out height);

            var unit = SelectedUnit;
            var widthPt = UnitConversion.ConvertToPoints(width, unit);
            var leftPt = UnitConversion.ConvertToPoints(left, unit);
            var topPt = UnitConversion.ConvertToPoints(top, unit);
            var lockAspect = lockAspectCheckBox.Checked;

            try
            {
                var app = Globals.ThisAddIn.Application;
                var sheet = (Excel.Worksheet)app.ActiveSheet;
                var shape = sheet.Shapes.Item(name);
                shape.LockAspectRatio = lockAspect ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                shape.Left = (float)leftPt;
                shape.Top = (float)topPt;
                shape.Width = (float)widthPt;
                if (!lockAspect && hasHeight)
                {
                    shape.Height = (float)UnitConversion.ConvertToPoints(height, unit);
                }

                MessageBox.Show(this, "画像のサイズ・位置を更新しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshShapeList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyPreset(int sizeMm)
        {
            var name = shapeListComboBox.SelectedItem as string;
            if (name == null)
            {
                MessageBox.Show(this, "編集する画像を選択してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var sizePt = (float)UnitConversion.MmToPoints(sizeMm);
                var app = Globals.ThisAddIn.Application;
                var sheet = (Excel.Worksheet)app.ActiveSheet;
                var shape = sheet.Shapes.Item(name);
                shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                shape.Width = sizePt;
                shape.Height = sizePt;

                MessageBox.Show(this, sizeMm + "mm角に変更しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshShapeList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
