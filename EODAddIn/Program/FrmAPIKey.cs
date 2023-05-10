using EODAddIn.Utils;
using Newtonsoft.Json;
using System;
using System.Security.Policy;
using System.Windows.Forms;

namespace EODAddIn.Program
{
    public partial class FrmAPIKey : Form
    {
        public FrmAPIKey()
        {
            InitializeComponent();
            txtAPI.Text = Program.APIKey;
        }

        /// <summary>
        /// Кнопка сохранить
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Save(object sender, EventArgs e)
        {
            try
            {
                EOD.Model.User user = JsonConvert.DeserializeObject<EOD.Model.User>(Response.GET("https://eodhistoricaldata.com/api/user", "api_token=" + txtAPI.Text));
                Program.SaveAPI(txtAPI.Text);
                MessageBox.Show("API key saved", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
            }
            catch (APIException ex)
            {
                if (ex.Code == 401)
                {
                    MessageBox.Show("Invalid API key", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        /// <summary>
        /// Кнопка отмены
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Cancel(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Кнопка перехода к прайс листу
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Price(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Program.UrlPrice);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(Program.UrlRegister);
        }
    }
}
