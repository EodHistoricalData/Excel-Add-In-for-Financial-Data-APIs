using EODAddIn.Program;
using EODAddIn.Utils;
using System;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmGetEtf : Form
    {
        public Model.FundamentalData Results;
        public string Tiker;

        public FrmGetEtf()
        {
            InitializeComponent();
            txtCode.Text = Settings.SettingsFields.EtfTicker;
        }

        private bool CheckForm()
        {
            if (string.IsNullOrEmpty(txtCode.Text))
            {
                MessageBox.Show("Insert a tiсker", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (!txtCode.Text.Contains("."))
            {
                MessageBox.Show("Insert exchange", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return;
            Tiker = txtCode.Text;

            try
            {
                Results = APIEOD.GetFundamental(Tiker);
                if (Results.ETF_Data == null)
                    throw new NullReferenceException("No ETF data");
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (Exception ex)
            {
                ErrorReport error = new ErrorReport(ex);
                error.ShowAndSend();
                return;
            }

            Settings.SettingsFields.EtfTicker = Tiker;
            Settings.Save();
            Close();
        }

        /// <summary>
        /// Отображение формы поиска тикера
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TsmiFindTicker_Click(object sender, EventArgs e)
        {
            FrmSearchTiker frm = new FrmSearchTiker();
            frm.ShowDialog();

            if (frm.Result.Code == null) return;

            txtCode.Text = $"{frm.Result.Code}.{frm.Result.Exchange}";
        }
    }
}
