using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmGetFundamental : Form
    {
        public Model.FundamentalData Results;
        public string Tiker;
        public string Exchange;
        public string Period;
        public DateTime From;
        public DateTime To;

        public FrmGetFundamental()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return; 

            Tiker = txtCode.Text;
            Exchange = txtExchange.Text;

            Results = Utils.APIEOD.GetFundamental($"{Tiker}.{Exchange}");

            Close();
        }

        private bool CheckForm()
        {
            if (string.IsNullOrEmpty(txtCode.Text))
            {
                MessageBox.Show("Insert a tiсker", "Error", MessageBoxButtons.OK,MessageBoxIcon.Error);
                return false;
            }
            if (string.IsNullOrEmpty(txtExchange.Text))
            {
                MessageBox.Show("Insert exchange", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
           
            return true;
        }
    }
}
