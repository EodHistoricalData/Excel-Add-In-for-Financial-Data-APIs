using EOD.Model.OptionsData;
using EOD.Model.Screener;

using EODAddIn.Forms;
using EODAddIn.Program;
using EODAddIn.View.Forms;

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;

using static EOD.API;

namespace EODAddIn.BL.Screener
{
    public class ScreenerManager
    {
        public List<Screener> Screeners { get; private set; }

        public ScreenerManager()
        {
            Screeners = Settings.Data.Screeners;
        }

        public async Task AddNewScreener()
        {
            if (Screeners.Count == 0)
            {
                FrmScreener frmScreener = new FrmScreener();
                FormShower.FrmShowDialog(frmScreener);

                if (frmScreener.DialogResult == DialogResult.OK)
                {
                    Screeners.Add(frmScreener.Screener);
                    Save();
                    await LoadAndPrintScreener(frmScreener.Screener);
                }
            }
            else
            {
                FrmScreenerDispatcher frmScreenerDispatcher = new FrmScreenerDispatcher(this);
                FormShower.FrmShowDialog(frmScreenerDispatcher);
            }
        }

        public void Save()
        {
            Settings.Data.Screeners = Screeners;
            Settings.Save();
        }

        public async Task<StockMarkerScreener> LoadData(Screener screener)
        {
            return await ScreenerAPI.GetScreener(GetFilters(screener), screener.Signals, screener.Sort, screener.Limit);
        }

        public async Task LoadAndPrintScreener(Screener screener)
        {

            var data = await LoadData(screener);
            ScreenerPrinter.PrintScreener(screener.NameScreener, data);
        }

        private List<(Field, Operation, string)> GetFilters(Screener screener)
        {
            List<(Field, Operation, string)> res = new List<(Field, Operation, string)>();
            foreach (Filter filter in screener.Filters)
            {
                res.Add((filter.Field, filter.Operation, filter.Value));
            }

            //if (!string.IsNullOrEmpty(screener.Code))
            //{
            //    (Field, Operation, string) newfilter = (Field.Code, Operation.Equals, screener.Code);
            //    res.Add(newfilter);
            //}
            //if (!string.IsNullOrEmpty(screener.Name))
            //{
            //    (Field, Operation, string) newfilter = (Field.Name, Operation.Equals, screener.Name);
            //    res.Add(newfilter);
            //}
            //if (!string.IsNullOrEmpty(screener.Exchange))
            //{
            //    (Field, Operation, string) filter = (Field.Exchange, Operation.Equals, screener.Exchange);
            //    res.Add(filter);
            //}
            //if (!string.IsNullOrEmpty(screener.Sector))
            //{
            //    (Field, Operation, string) filter = (Field.Sector, Operation.Equals, screener.Sector);
            //    res.Add(filter);
            //}
            //if (!string.IsNullOrEmpty(screener.Industry))
            //{
            //    (Field, Operation, string) filter = (Field.Industry, Operation.Equals, screener.Industry);
            //    res.Add(filter);
            //}

            return res;
        }
    }
}
