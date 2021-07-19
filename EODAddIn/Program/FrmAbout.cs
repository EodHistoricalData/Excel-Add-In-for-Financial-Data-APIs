using System;
using System.Windows.Forms;

namespace EODAddIn.Program
{
    public partial class FrmAbout : Form
    {
        public FrmAbout()
        {
            InitializeComponent();
            lblVersion.Text = Program.Version.Text;
            lnkSite.Text = Program.UrlCompany;
            lblProgramName.Text += Program.ProgramName;
        }

        private void LnkSite_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(Program.UrlCompany);
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
