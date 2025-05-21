using System;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace ProJetsAddin
{
    public partial class ThisAddIn
    {
        private MainForm mainForm;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Initialize the main form
            mainForm = new MainForm();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (mainForm != null)
            {
                mainForm.Close();
                mainForm.Dispose();
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return Globals.Factory.GetRibbonFactory().CreateRibbon<ProJetsRibbon>();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
} 