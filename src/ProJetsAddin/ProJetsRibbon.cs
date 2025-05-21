using System;
using Microsoft.Office.Tools.Ribbon;

namespace ProJetsAddin
{
    public partial class ProJetsRibbon
    {
        private void ProJetsRibbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        private void ShowProJetsButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.mainForm != null)
            {
                Globals.ThisAddIn.mainForm.Show();
                Globals.ThisAddIn.mainForm.Activate();
            }
        }
    }
} 