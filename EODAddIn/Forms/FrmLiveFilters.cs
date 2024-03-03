using EODAddIn.BL.Live;

using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmLiveFilters : Form, IDisposable
    {
        internal List<Filter> Filters;
        public FrmLiveFilters()
        {
            InitializeComponent();
        }

        public FrmLiveFilters(List<Filter> filters)
        {
            InitializeComponent();

            Filters = filters;

            for (int i = 0; i < filters.Count; i++)
            {
                ClbFilters.SetItemChecked(i, filters[i].IsChecked);
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

        private List<Filter> SaveFilters()
        {
            List<Filter> filters = new List<Filter>();
            for (int i = 0; i < Filters.Count; i++)
            {
                var filter = new Filter()
                {
                    Name = ClbFilters.GetItemText(ClbFilters.Items[i]), 
                    IsChecked = ClbFilters.GetItemChecked(i)
                };
                filters.Add(filter);
            }
            return filters;
        }
    }
}
