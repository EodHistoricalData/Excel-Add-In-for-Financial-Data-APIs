using EODAddIn.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Security.Policy;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmSearchTiker : Form
    {
        public EOD.Model.SearchResult Result = new EOD.Model.SearchResult();
        public FrmSearchTiker()
        {
            InitializeComponent();
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

            Close();
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSearch.Text)) return;
            gridResult.Rows.Clear();
            List<EOD.Model.SearchResult> searchResults = JsonConvert.DeserializeObject<List<EOD.Model.SearchResult>>(Response.GET($"https://eodhd.com/api/search/{txtSearch.Text}?api_token={Program.Program.APIKey}&fmt=json"));

            foreach (EOD.Model.SearchResult result in searchResults)
            {

                int i = gridResult.Rows.Add();

                gridResult.Rows[i].Cells[0].Value = result.Code;
                gridResult.Rows[i].Cells[1].Value = result.Exchange;
                gridResult.Rows[i].Cells[2].Value = result.Name;

            }
        }

        private void GridResult_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Result.Code = gridResult.Rows[e.RowIndex].Cells[0].Value.ToString();
            Result.Exchange = gridResult.Rows[e.RowIndex].Cells[1].Value.ToString();
            Result.Name = gridResult.Rows[e.RowIndex].Cells[2].Value.ToString();

            Close();
        }
    }
}
