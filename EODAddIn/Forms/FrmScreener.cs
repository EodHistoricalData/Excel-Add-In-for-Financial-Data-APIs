using System;
using System.Collections.Generic;
using System.Windows.Forms;
using static EOD.API;

namespace EODAddIn.Forms
{
    public partial class FrmScreener : Form
    {
        public List<(Field, Operation, string)> Filters { get; set; } = new List<(Field, Operation, string)>();
        public string Signals { get; set; }
        public string Sort { get; set; }
        public int Limit { get; set; }
        
        private Dictionary<string, string> _fields = new Dictionary<string, string>()
        {
            { "code", "string" },
            { "name", "string" },
            { "exchange", "string" },
            { "sector", "string" },
            { "industry", "string" },
            { "market_capitalization", "number" },
            { "earnings_share", "number" },
            { "dividend_yield", "number" },
            { "refund_1d_p", "number" },
            { "refund_5d_p", "number" },
            { "avgvol_1d", "number" },
            { "avgvol_200d", "number" },
        };
        private List<string> _operationNumber = new List<string>()
        {
            { "=" },
            { ">" },
            { "<" },
            { ">=" },
            { "<=" },
            { "!=" },
        };
        private List<string> _operationString = new List<string>()
        {
            { "=" },
        };

        public FrmScreener()
        {
            InitializeComponent();
        }

        private void dataGridViewFilters_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)dataGridViewFilters.Rows[e.RowIndex].Cells[colField.Index];
            foreach (var item in _fields)
            {
                cell.Items.Add(item.Key);
            }
        }

        private void dataGridViewFilters_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == colField.Index && e.RowIndex > -1)
            {
                string val = dataGridViewFilters.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                if (_fields.ContainsKey(val))
                {
                    DataGridViewComboBoxCell cell;
                    List<string> lst;
                    if (_fields[val] == "string")
                    {
                        cell = (DataGridViewComboBoxCell)dataGridViewFilters.Rows[e.RowIndex].Cells[colOperation.Index];
                        lst = _operationString;                   
                    }
                    else
                    {
                        cell = (DataGridViewComboBoxCell)dataGridViewFilters.Rows[e.RowIndex].Cells[colOperation.Index];
                        lst = _operationNumber;
                    }

                    cell.Items.Clear();
                    foreach (var item in lst)
                    {
                        cell.Items.Add(item);
                    }

                    if (cell.Items.Count == 1)
                    {
                        cell.Value = cell.Items[0].ToString();
                    }
                }
            }
        }

        private void btnAddFilter_Click(object sender, EventArgs e)
        {
            dataGridViewFilters.Rows.Add();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                SetFilteres();
                SetSignals();
                SetSort();
                Limit = (int)numLimit.Value;
                DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Screener error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void SetFilteres()
        {
            Filters.Clear();
            foreach (DataGridViewRow row in dataGridViewFilters.Rows)
            {
                if (row.Cells[colField.Index].Value == null) continue;
                if (row.Cells[colOperation.Index].Value == null) throw new Exception("Select operation type from the list");
                if (row.Cells[colValue.Index].Value == null) throw new Exception("Select a value for the operation");

                Field field;
                Operation operation;

                switch (row.Cells[colField.Index].Value.ToString())
                {
                    case "code":
                        field = Field.Code;
                        break;

                    case "name":
                        field = Field.Name;
                        break;
                    case "exchange":
                        field = Field.Exchange;
                        break;
                    case "sector":
                        field = Field.Sector;
                        break;
                    case "industry":
                        field = Field.Industry;
                        break;
                    case "market_capitalization":
                        field = Field.MarketCapitalization;
                        break;
                    case "earnings_share":
                        field = Field.EarningsShare;
                        break;
                    case "dividend_yield":
                        field = Field.DividendYield;
                        break;
                    case "refund_1d_p":
                        field = Field.Refund1dP;
                        break;
                    case "refund_5d_p":
                        field = Field.Refund5dP;
                        break;
                    case "avgvol_1d":
                        field = Field.Refund5dP;
                        break;
                    case "avgvol_200d":
                        field = Field.Refund5dP;
                        break;

                    default:
                        throw new Exception("Select a field");
                }

                switch (row.Cells[colOperation.Index].Value.ToString())
                {
                    case "=":
                        operation = Operation.Equals;
                        break;
                    case ">":
                        operation = Operation.More;
                        break;
                    case "<":
                        operation = Operation.Less;
                        break;
                    case ">=":
                        operation = Operation.NotLess;
                        break;
                    case "<=":
                        operation = Operation.NotMore;
                        break;
                    case "!=":
                        operation = Operation.NotEquals;
                        break;
                    
                    default:
                        throw new Exception("Select a operation");
                }

                (Field, Operation, string) filter = (field, operation, row.Cells[colValue.Index].Value.ToString());
                Filters.Add(filter);
            }
        }

        private void SetSignals()
        {
            Signals = string.Empty;
            if (chk50d_new_lo.Checked) Signals += "50d_new_lo,";
            if (chk50d_new_hi.Checked) Signals += "50d_new_hi,";
            if (chk200d_new_lo.Checked) Signals += "200d_new_lo,";
            if (chk200d_new_hi.Checked) Signals += "200d_new_hi,";
            if (chkBookvalue_neg.Checked) Signals += "bookvalue_neg,";
            if (chkBookvalue_pos.Checked) Signals += "bookvalue_pos,";
            if (chkWallstreet_lo.Checked) Signals += "wallstreet_lo,";
            if (chkWallstreet_hi.Checked) Signals += "wallstreet_hi,";

            if (Signals.Length > 0) Signals = Signals.Substring(0, Signals.Length - 1);
        }

        private void SetSort()
        {
            Sort = string.Empty;

            if (cboSortField.SelectedIndex == -1) return;
            Sort = cboSortField.Text;

            if (rbtnSortAsc.Checked)
            {
                Sort += ".asc";
            }
            else
            {
                Sort += ".desc";
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
