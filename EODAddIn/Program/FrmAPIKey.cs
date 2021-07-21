using EODAddIn.Utils;

using System;
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
                Model.User user = APIEOD.User(txtAPI.Text);
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
        /// Кнопка регистрации
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Register(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Program.UrlCompany);
        }

        /// <summary>
        /// Кнопка перехода к API ключу
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CopyAPI(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Program.UrlKey);
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
    }
}
