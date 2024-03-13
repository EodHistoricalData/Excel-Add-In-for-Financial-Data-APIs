using System;

using static EOD.API;

namespace EODAddIn.BL.Screener
{
    [Serializable]
    public class Filter
    {
        public Field Field { get; set; }
        public Operation Operation { get; set; }
        public string Value { get; set; }
    }
}
