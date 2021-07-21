using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            Program.SaveAPI(txtAPI.Text);
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
