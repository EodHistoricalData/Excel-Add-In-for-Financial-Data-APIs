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
            gridResult.Rows.Clear();
            List<SearchResult> searchResults = APIEOD.Search(txtSearch.Text);

            foreach (SearchResult result in searchResults)
            {

                int i = gridResult.Rows.Add();

                gridResult.Rows[i].Cells[0].Value = result.Code;
                gridResult.Rows[i].Cells[1].Value = result.Exchange;
                gridResult.Rows[i].Cells[2].Value = result.Name;

            }
        }

        private void GridResult_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Result.Code = gridResult.SelectedRows[e.RowIndex].Cells[0].Value.ToString();
            Result.Exchange = gridResult.SelectedRows[e.RowIndex].Cells[1].Value.ToString();
            Result.Name = gridResult.SelectedRows[e.RowIndex].Cells[2].Value.ToString();

            Close();
        }
    }
}
