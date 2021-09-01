using EODAddIn.Model;
using EODAddIn.Utils;

using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmSearchTiker : Form
    {
        public SearchResult Result = new SearchResult();
        public FrmSearchTiker()
        {
            InitializeComponent();
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            List<SearchResult> searchResults = APIEOD.Search(txtSearch.Text);

            foreach (SearchResult result in searchResults)
            {

                int i = gridResult.Rows.Add();

                gridResult.Rows[i].Cells[0].Value = result.Code;
                gridResult.Rows[i].Cells[1].Value = result.Exchange;
                gridResult.Rows[i].Cells[2].Value = result.Name;
                gridResult.Rows[i].Cells[3].Value = result.Country;
                gridResult.Rows[i].Cells[4].Value = result.Currency;
            }
        }

        private void BtnSelect_Click(object sender, EventArgs e)
        {
            if (gridResult.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select ticker", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Result.Code = gridResult.SelectedRows[0].Cells[0].Value.ToString();
            Result.Exchange = gridResult.SelectedRows[0].Cells[1].Value.ToString();
            Result.Name = gridResult.SelectedRows[0].Cells[2].Value.ToString();
            Result.Country = gridResult.SelectedRows[0].Cells[3].Value.ToString();
            Result.Currency = gridResult.SelectedRows[0].Cells[4].Value.ToString();

            Close();
        }
    }
}
