using EOD.Model.Screener;

using Microsoft.Office.Core;

using System;
using System.Collections.Generic;

using static EOD.API;

namespace EODAddIn.BL.Screener
{
    [Serializable]
    public class Screener
    {
        public string NameScreener { get; set; }
        public string Sector { get; set; }
        public string Industry { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public string Exchange { get; set; }
        public int Limit { get; set; } = 100;

        //public bool Screener200d_New_Hi { get; set; }
        //public bool Screener200d_New_Lo { get; set; }
        //public bool ScreenerBookValue_Neg { get; set; }
        //public bool ScreenerBookValue_Pos { get; set; }
        //public bool ScreenerWallStreet_Lo { get; set; }
        //public bool ScreenerWallStreet_Hi { get; set; }
        //public bool ScreenerRbtnSortAsc { get; set; }
        //public bool ScreenerRbtnSortDesc { get; set; }

        public List<Filter> Filters { get; set; } = new List<Filter>();
        public List<Signal> Signals { get; set; } = null;
        //{
        //    get
        //    {
        //        var res = new List<Signal>();

        //        if (Screener200d_New_Lo) res.Add(Signal.New_200d_low);
        //        if (Screener200d_New_Hi) res.Add(Signal.New_200d_hi);
        //        if (ScreenerBookValue_Neg) res.Add(Signal.Bookvalue_neg);
        //        if (ScreenerBookValue_Pos) res.Add(Signal.Bookvalue_pos);
        //        if (ScreenerWallStreet_Lo) res.Add(Signal.Wallstreet_low);
        //        if (ScreenerWallStreet_Hi) res.Add(Signal.Wallstreet_hi);

        //        if (res.Count == 0) return null;
        //        return res;
        //    }
        //}
        public (Field, Order)? Sort { get; set; }


        public Screener()
        {

        }

    }
}
