using System.Drawing;
using System.Windows.Forms;

namespace SealAddIn.TaskPane
{
    /// <summary>タスクペインのルートコントロール。4つの機能タブをホストします。</summary>
    public class MainTaskPaneControl : UserControl
    {
        private readonly ShapeEditorControl shapeEditorControl;
        private readonly StampHistoryControl stampHistoryControl;
        private readonly TabControl tabControl;

        public MainTaskPaneControl()
        {
            Dock = DockStyle.Fill;
            // Officeのグレー系/ダークテーマ配下でも文字が読めるよう、テーマに依存しない配色を明示する。
            BackColor = Color.White;
            ForeColor = Color.Black;

            var imageLibraryControl = new ImageLibraryControl { Dock = DockStyle.Fill };
            shapeEditorControl = new ShapeEditorControl { Dock = DockStyle.Fill };
            var stampGeneratorControl = new StampGeneratorControl { Dock = DockStyle.Fill };
            stampHistoryControl = new StampHistoryControl { Dock = DockStyle.Fill };

            tabControl = new TabControl { Dock = DockStyle.Fill, BackColor = Color.White, ForeColor = Color.Black };
            tabControl.TabPages.Add(BuildPage("画像ライブラリ", imageLibraryControl));
            tabControl.TabPages.Add(BuildPage("サイズ・位置編集", shapeEditorControl));
            tabControl.TabPages.Add(BuildPage("スタンプ作成", stampGeneratorControl));
            tabControl.TabPages.Add(BuildPage("押印履歴", stampHistoryControl));
            tabControl.SelectedIndexChanged += (s, e) => OnTabSelected();

            Controls.Add(tabControl);
        }

        private static TabPage BuildPage(string title, Control content)
        {
            var page = new TabPage(title) { BackColor = Color.White, ForeColor = Color.Black };
            page.Controls.Add(content);
            return page;
        }

        private void OnTabSelected()
        {
            if (tabControl.SelectedTab == null)
            {
                return;
            }

            switch (tabControl.SelectedTab.Text)
            {
                case "サイズ・位置編集":
                    shapeEditorControl.RefreshShapeList();
                    break;
                case "押印履歴":
                    stampHistoryControl.RefreshList();
                    break;
            }
        }
    }
}
