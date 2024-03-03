using EOD.Model;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;


using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using System.Windows.Threading;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

using static EODAddIn.Utils.ExcelUtils;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;

namespace EODAddIn.BL.Live
{
    [Serializable]
    public class LiveDownloader
    {

        public Guid Id { get; set; }
        public string Name { get; set; }
        public List<Ticker> Tickers { get; set; }
        public List<Filter> Filters { get; set; }
        public int Interval { get; set; }
        public bool Smart { get; set; }

        [XmlIgnore]
        public bool? IsActive
        {
            get => _isActivate;
            private set
            {
                _isActivate = value;
                ActiveChanged?.Invoke(this, EventArgs.Empty);
            }
        }
        private bool? _isActivate;

        public event EventHandler ActiveChanged;
        private CustomXMLPart _customXMLPart;

        public event EventHandler OnDeleted;

        public List<ExchangeDownloadRules> Rules { get; set; } = new List<ExchangeDownloadRules>();

        private CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();
        [XmlIgnore]
        public Workbook Workbook
        {
            get;
            private set;
        }

        private static readonly SemaphoreSlim _semaphoreSlim = new SemaphoreSlim(1, 1);

        private EOD.API API { get; set; } = new EOD.API(Program.Program.APIKey);

        /// <summary>
        /// For serialize
        /// </summary>
        public LiveDownloader()
        {

        }

        public LiveDownloader(List<Ticker> tickers, int interval, bool smart, List<Filter> filters, string name, Workbook workbook)
        {
            Tickers = tickers;
            Filters = filters;
            Interval = interval;
            Id = Guid.NewGuid();

            Smart = smart;
            Name = name;
            Workbook = workbook;
            Workbook.BeforeClose += Workbook_BeforeClose;

            Dispatcher dispUI = Dispatcher.CurrentDispatcher;
            dispUI.Invoke(CreateRules);

        }

        public void Set(Workbook workbook, CustomXMLPart customXMLPart)
        {
            Workbook = workbook;
            _customXMLPart = customXMLPart;

        }

        private void Workbook_BeforeClose(ref bool Cancel)
        {
            Stop();
        }

        public void Save()
        {

            XmlSerializer xmlSerializer = new XmlSerializer(typeof(LiveDownloader));
            var xml = "";
            using (var sww = new StringWriter())
            {
                using (XmlWriter writer = XmlWriter.Create(sww))
                {
                    xmlSerializer.Serialize(writer, this);
                    xml = sww.ToString();
                }
            }

            _customXMLPart?.Delete();
            _customXMLPart = Workbook.CustomXMLParts.Add(xml);
        }

        public async void Start()
        {
            _cancellationTokenSource.Cancel();
            _cancellationTokenSource = new CancellationTokenSource();
            await DoWork(_cancellationTokenSource.Token);

        }

        public void Stop()
        {
            _cancellationTokenSource.Cancel();
        }

        public void Delete()
        {
            _cancellationTokenSource.Cancel();
            _customXMLPart?.Delete();
            _customXMLPart = null;
            OnDeleted?.Invoke(this, EventArgs.Empty);
        }

        private async void CreateRules()
        {
            EOD.API api = new EOD.API(Program.Program.APIKey);
            var exchanges = Tickers.GroupBy(x => x.Exchange);
            foreach (var exchange in exchanges)
            {
                try
                {
                    var details = await api.GetExchangeDetailsAsync(exchange.Key);
                    var rules = new ExchangeDownloadRules(details);
                    Rules.Add(rules);
                }
                catch
                {
                    throw new Exception("It was not possible to download an exchange working hours for one or more tickers. Please check that the ticker list is correct.");
                }
            }
        }

        public string GetTickers()
        {
            string result = "";
            foreach (var ticker in Tickers)
            {
                result += ticker.FullName + ", ";
            }

            if (result.Length < 3) return string.Empty;
            return result.Remove(result.LastIndexOf(','), 2);
        }

        private async Task DoWork(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    IsActive = true;
                    List<Ticker> tickers = GetActiveTickers();
                    var excenges = tickers.GroupBy(x => x.Exchange).Select(x => x.Key).ToList();
                    if (tickers.Count == 0)
                    {
                        Stop();
                        return;
                    }

                    List<LiveStockPrice> list = new List<LiveStockPrice>();
                    if (excenges.Count() == 1)
                    {
                        foreach (var ticker in tickers)
                        {
                            var result = await API.GetLiveStockPricesAsync(ticker.FullName);
                            list.Add(result);
                        }
                    }
                    else
                    {
                        foreach (var ticker in tickers)
                        {
                            var result = await API.GetLiveStockPricesAsync(ticker.FullName, excenges);
                            list.AddRange(result);
                        }
                    }
                    try
                    {
                        await _semaphoreSlim.WaitAsync();
                        PrintLive(list);
                    }
                    catch { }
                    finally
                    {
                        _semaphoreSlim.Release();
                    }

                    await Task.Delay(TimeSpan.FromMinutes(Interval), token);
                }
                catch
                {
                    IsActive = false;
                }
            }
        }

        private List<Ticker> GetActiveTickers()
        {
            List<Ticker> tickers = new List<Ticker>();

            var exchanges = Tickers.GroupBy(x => x.Exchange);
            foreach (var exchange in exchanges)
            {
                if (CheckWorkTime(exchange.Key))
                {
                    foreach (var code in exchange)
                    {
                        tickers.Add(code);
                    }
                }
                else
                {
                    foreach (var code in exchange)
                    {
                        tickers.Add(code);
                    }
                }
            }
            return tickers;
        }

        /// <summary>
        /// Collect WSs for tickers data
        /// </summary>
        /// <param name="data"></param>
        /// <param name="tickers"></param>
        private void PrintLive(List<LiveStockPrice> data)
        {
            try
            {
                OnStart();
                SetNonInteractive();

                Worksheet sh = null;
                foreach (Worksheet ws in Workbook.Worksheets)
                {
                    if (ws.Name == Name)
                    {
                        sh = ws;
                    }
                }
                if (sh == null)
                {
                    sh = AddSheet(Name);
                }

                PrintData(data, sh);
            }
            catch
            {

            }
            finally
            {
                SetInteractive();
                OnEnd();
            }
        }

        /// <summary>
        /// Identify the sheet(s) and print data
        /// </summary>
        /// <param name="data"></param>
        /// <param name="targetWSs">ticker - Worksheet</param>
        private void PrintData(List<LiveStockPrice> data, Worksheet worksheet)
        {
            Range usedRange = worksheet.UsedRange;
            int rows = usedRange.Rows.Count;
            if (rows > 1)
            {
                //old - need to seek for row
                foreach (var dataRow in data)
                {
                    bool found = false;
                    for (int row = 2; row <= usedRange.Rows.Count; row++)
                    {
                        if (worksheet.Cells[row, 1].Value == dataRow.Code)
                        {
                            found = true;
                            for (int col = 2; col <= usedRange.Columns.Count; col++)
                            {
                                if (worksheet.Cells[1, col].Value != null)
                                {
                                    worksheet.Cells[row, col] = dataRow.GetType().GetProperty(worksheet.Cells[1, col].Value).GetValue(dataRow, null);
                                }
                            }
                        }
                    }
                    if (!found)
                    {
                        usedRange = worksheet.UsedRange;
                        worksheet.Cells[usedRange.Rows.Count + 1, 1] = dataRow.Code;
                        for (int col = 2; col <= usedRange.Columns.Count; col++)
                        {
                            if (worksheet.Cells[1, col].Value != null)
                            {
                                worksheet.Cells[usedRange.Rows.Count + 1, col] = dataRow.GetType().GetProperty(worksheet.Cells[1, col].Value).GetValue(dataRow, null);
                            }
                        }
                    }
                }
            }
            else
            {
                //new
                var props = Filters.Where(x => x.IsChecked == true).Select(x => x.Name).ToList();
                if (props.Contains("Code"))
                {
                    string item = props[props.IndexOf("Code")];
                    props.RemoveAt(props.IndexOf("Code"));
                    props.Insert(0, item);
                }
                else
                {
                    props.Insert(0, "Code");
                }
                object[,] table = new object[data.Count + 1, props.Count];
                for (int j = 0; j < props.Count; j++)
                {
                    table[0, j] = props[j];
                }
                for (int i = 0; i < data.Count; i++)
                {
                    foreach (var prop in props)
                    {
                        table[i + 1, props.IndexOf(prop)] = data[i].GetType().GetProperty(prop).GetValue(data[i], null);
                    }
                }
                Range tableRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[table.GetLength(0), table.GetLength(1)]];
                tableRange.Value = table;

                if (Smart)
                {
                    MakeTable(tableRange, worksheet, Name, 1);
                }
            }


        }

        /// <summary>
        /// Check if stock market is open now
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        private bool CheckWorkTime(string key)
        {
            var rule = Rules.Find(x => x.Exchange == key.ToUpper());
            DateTime utc = DateTime.UtcNow;
            if (rule == null)
            {
                return true;
            }
            DateTime stockNow = utc.AddHours(rule.Gtmoffset);
            bool isHoliday = rule.Holidays.Contains(stockNow.Date);
            bool isWorkDay = rule.DaysofWeek.Contains(stockNow.ToString("ddd", new CultureInfo("en-GB")));
            bool isWorkHour = stockNow.TimeOfDay >= rule.Open.TimeOfDay && stockNow.TimeOfDay <= rule.Close.TimeOfDay;
            return !isHoliday && isWorkDay && isWorkHour;
        }
    }
}
