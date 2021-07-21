using Microsoft.Office.Tools.Ribbon;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EODAddIn
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BtnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            Program.FrmAbout frm = new Program.FrmAbout();
            frm.ShowDialog();
        }

        private void BtnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            Program.FrmAPIKey frm = new Program.FrmAPIKey();
            frm.ShowDialog();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
           // Model.User user = Utils.EODAPI.User();
           // user.Email = "asd";
            //Utils.EODAPI.User();
        }
    }
}
