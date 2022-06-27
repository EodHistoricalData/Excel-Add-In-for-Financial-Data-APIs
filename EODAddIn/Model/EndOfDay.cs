using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class EndOfDay
    {
        public DateTime? Date { get; set; }
        public double? Open { get; set; }
        public double? High { get; set; }
        public double? Low { get; set; }
        public double? Close { get; set; }
        public double? Adjusted_close { get; set; }

        public double? Adjusted_low
        {
            get
            {
                if (_k == null) SetK();

                return Low * _k;
            }
        }

        public double? Adjusted_high
        {
            get
            {
                if (_k == null) SetK();

                return High * _k;
            }
        }

        public double? Adjusted_open
        {
            get
            {
                if (_k == null) SetK();

                return Open * _k;
            }
        }

        public long? Volume { get; set; }

        private double? _k;

        private void SetK ()
        {
            _k = Adjusted_close / Close;
        }

    }
}
