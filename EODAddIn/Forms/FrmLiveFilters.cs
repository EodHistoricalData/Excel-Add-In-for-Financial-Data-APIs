using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmLiveFilters : Form, IDisposable
    {
        internal List<(string, bool)> Filters;
        public FrmLiveFilters()
        {
            InitializeComponent();
        }

        public FrmLiveFilters(List<(string, bool)> filters)
        {
            InitializeComponent();

            Filters = filters;

            for (int i = 0; i < filters.Count; i++)
            {
                ClbFilters.SetItemChecked(i, filters[i].Item2);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            Filters = SaveFilters();
            DialogResult = DialogResult.OK;
        }

        private List<(string, bool)> SaveFilters()
        {
            List<(string, bool)> filters = new List<(string, bool)>();
            for (int i = 0; i < Filters.Count; i++)
            {
                filters.Add((ClbFilters.GetItemText(ClbFilters.Items[i]), ClbFilters.GetItemChecked(i)));
            }
            return filters;
        }
    }
}
