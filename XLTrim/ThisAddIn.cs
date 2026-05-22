using System;
using Microsoft.Office.Core;
using XLTrim.Ribbon;

namespace XLTrim
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e) { }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new XLTrimRibbon();
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
