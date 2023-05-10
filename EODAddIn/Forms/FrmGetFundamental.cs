using EODAddIn.BL.FundamentalDataAPI;
using EODAddIn.Program;
using EODAddIn.Utils;

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
        public EOD.Model.Fundamental.FundamentalData Results;
        public string Tiker;
        public string Exchange;
        public string Period;
        public DateTime From;
        public DateTime To;

        public FrmGetFundamental()
        {
            InitializeComponent();
            txtCode.Text = Program.Settings.SettingsFields.FundamentalTicker;
        }

        /// <summary>
        /// Кнопка загрузки данных
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return; 

            Tiker = txtCode.Text;

            try
            {
                Results = FundamentalDataAPI.GetFundamental(Tiker);
                DialogResult = DialogResult.OK;
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
            
            Program.Settings.SettingsFields.FundamentalTicker = Tiker;
            Program.Settings.Save();
            Close();
        }

        /// <summary>
        /// Проверка корректности заполнения формы
        /// </summary>
        /// <returns></returns>
        private bool CheckForm()
        {
            if (string.IsNullOrEmpty(txtCode.Text))
            {
                MessageBox.Show("Insert a tiсker", "Error", MessageBoxButtons.OK,MessageBoxIcon.Error);
                return false;
            }
            if (!txtCode.Text.Contains("."))
            {
                MessageBox.Show("Insert exchange", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
           
            return true;
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
