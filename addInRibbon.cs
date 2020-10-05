using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace mailBoxWizard
{
    public partial class addInRibbon
    {
        
        private void addInRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            DataView myform = new DataView();
            myform.Show();
        }
    }
}
