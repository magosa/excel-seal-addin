using System;
using Microsoft.Office.Tools.Excel;

namespace SealAddIn
{
    public partial class ThisAddIn
    {
        /// <summary>リボンのイベントハンドラーからタスクペインを操作するために保持する参照。</summary>
        internal Ribbon.StampRibbon Ribbon;

        private Microsoft.Office.Tools.CustomTaskPane stampTaskPane;
        private TaskPane.MainTaskPaneControl stampTaskPaneControl;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon.StampRibbon();
        }

        internal bool IsTaskPaneVisible()
        {
            return stampTaskPane != null && stampTaskPane.Visible;
        }

        internal void SetTaskPaneVisibility(bool visible)
        {
            EnsureTaskPane();
            stampTaskPane.Visible = visible;
        }

        internal void ShowOptionsDialog()
        {
            using (var form = new Options.OptionsForm())
            {
                form.ShowDialog();
            }
        }

        private void EnsureTaskPane()
        {
            if (stampTaskPane != null)
            {
                return;
            }

            stampTaskPaneControl = new TaskPane.MainTaskPaneControl();
            stampTaskPane = this.CustomTaskPanes.Add(stampTaskPaneControl, "画像配置・電子押印ツール");
            stampTaskPane.Width = 380;
            stampTaskPane.VisibleChanged += (s, e) => Ribbon?.InvalidateTaskPaneButton();
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
