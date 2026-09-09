using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace SealAddIn.Ribbon
{
    [ComVisible(true)]
    public class StampRibbon : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("SealAddIn.Ribbon.StampRibbon.xml"))
            using (var reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public void OnLoad(IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
            Globals.ThisAddIn.Ribbon = this;
        }

        public bool GetTaskPanePressed(IRibbonControl control)
        {
            return Globals.ThisAddIn.IsTaskPaneVisible();
        }

        public void OnToggleTaskPane(IRibbonControl control, bool pressed)
        {
            Globals.ThisAddIn.SetTaskPaneVisibility(pressed);
        }

        public void OnOpenOptions(IRibbonControl control)
        {
            Globals.ThisAddIn.ShowOptionsDialog();
        }

        public void InvalidateTaskPaneButton()
        {
            ribbon?.InvalidateControl("ToggleTaskPaneButton");
        }
    }
}
