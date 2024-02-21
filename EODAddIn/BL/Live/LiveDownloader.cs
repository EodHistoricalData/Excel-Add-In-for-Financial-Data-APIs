using EOD.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WinForms = System.Windows.Forms;
using static EODAddIn.Utils.ExcelUtils;
using System.Windows.Threading;

namespace EODAddIn.BL.Live
{
    public class LiveDownloader
    {
        public List<(string, string)> Tickers { get; set; }
        public List<(string, bool)> Filters { get; set; }
        /// <summary>
        /// ticker - Worksheet Name
        /// </summary>
        public List<(string, string)> WsNames { get; set; } = new List<(string, string)>();
        public int Interval { get; set; }
        public int Output { get; set; }
        public bool Smart { get; set; }
        public List<ExchangeDownloadRules> Rules { get; set; } = new List<ExchangeDownloadRules>();
        public string Name { get; set; }
        private static SemaphoreSlim SemaphoreSlim = new SemaphoreSlim(1, 1);

        public delegate void StatusHandler(object sender, EventArgs e);
        public event StatusHandler OnStatusChanged;
        public bool? IsActive
        {
            get => isActive;
            set
            {
                isActive = value;
                OnStatusChanged?.Invoke(this, EventArgs.Empty);
            }
        }
        private bool? isActive = false;

        private EOD.API API { get; set; } = new EOD.API(Program.Program.APIKey);
        public LiveDownloader()
        {

        }
        public LiveDownloader(List<(string, string)> tickers, int interval, int output, bool smart, List<(string, bool)> filters, string name, List<(string, string)> wsNames)
        {
            Tickers = tickers;

            Dispatcher dispUI = Dispatcher.CurrentDispatcher;
            dispUI.Invoke(CreateRules);

            Interval = interval;
            Output = output;
            Smart = smart;
            Filters = filters;
            Name = name;
            WsNames = wsNames;
        }

        private async void CreateRules()
        {
            EOD.API api = new EOD.API(Program.Program.APIKey);
            var exchanges = Tickers.GroupBy(x => x.Item2);
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
                result += ticker.Item1 + "." + ticker.Item2 + ", ";
            }
            return result.Remove(result.LastIndexOf(','), 2);
        }

        internal async Task RequestAndPrint(CancellationToken token)
        {
            bool first = true;
            while (true)
            {
                if (token.IsCancellationRequested) break;
                try
                {
                    IsActive = true;
                    List<string> tickers = new List<string>();
                    var exchanges = Tickers.GroupBy(x => x.Item2);
                    foreach (var exchange in exchanges)
                    {
                        if (CheckWorkTime(exchange.Key))
                        {
                            foreach (var code in exchange)
                            {
                                tickers.Add(code.Item1 + "." + code.Item2);
                            }
                        }
                        else
                        {
                            if (first)
                            {
                                foreach (var code in exchange)
                                {
                                    tickers.Add(code.Item1 + "." + code.Item2);
                                }
                            }
                            IsActive = null;
                        }
                    }
                    first = false;
                    if (tickers.Count != 0)
                    {
                        string firstTicker = tickers[0];
                        List<LiveStockPrice> list = new List<LiveStockPrice>();
                        if (tickers.Count == 1)
                        {
                            var result = await API.GetLiveStockPricesAsync(firstTicker);
                            list.Add(result);
                        }
                        else
                        {
                            var result = await API.GetLiveStockPricesAsync(firstTicker, tickers);
                            list = result;
                        }
                        try
                        {
                            await SemaphoreSlim.WaitAsync();
                            PrintLive(list, tickers);
                        }
                        catch { }
                        finally
                        {
                            SemaphoreSlim.Release();
                        }
                    }
                    await Task.Delay(TimeSpan.FromMinutes(Interval), token);
                }
                catch
                {
                    IsActive = false;
                }
            }
        }

        /// <summary>
        /// Collect WSs for tickers data
        /// </summary>
        /// <param name="data"></param>
        /// <param name="tickers"></param>
        private void PrintLive(List<LiveStockPrice> data, List<string> tickers)
        {
            try
            {
                OnStart();
                SetNonInteractive();
                Application excel = Globals.ThisAddIn.Application;
                Dictionary<string, Worksheet> targetWSs = new Dictionary<string, Worksheet>();
                if (Output == 0)
                {
                    //=1 ws
                    string wsName = WsNames.Find(x => x.Item1 == "").Item2;
                    Worksheet targetWS = null;
                    foreach (Worksheet ws in excel.Worksheets)
                    {
                        if (ws.Name == wsName)
                        {
                            targetWS = ws;
                        }
                    }
                    if (targetWS == null)
                    {
                        WsNames.Clear();
                        targetWS = AddSheet(Name);
                        WsNames.Add(("", targetWS.Name));
                    }
                    targetWSs.Add("", targetWS);
                }
                else
                {
                    //>1 ws
                    foreach (var ticker in tickers)
                    {
                        bool found = false;
                        string codeName = WsNames.Find(x => x.Item1 == ticker).Item2;
                        foreach (Worksheet ws in excel.Worksheets)
                        {
                            if (ws.Name == codeName)
                            {
                                targetWSs.Add(ticker, ws);
                                found = true;
                            }
                        }
                        if (!found)
                        {
                            Worksheet targetWS = AddSheet(Name + " " + ticker);
                            int index = WsNames.IndexOf(WsNames.Find(x => x.Item1 == ticker));
                            if (index > -1)
                            {
                                WsNames[index] = (ticker, targetWS.Name);
                            }
                            else
                            {
                                WsNames.Add((ticker, targetWS.Name));
                            }
                            targetWSs.Add(ticker, targetWS);
                        }
                    }
                }
                PrintData(data, targetWSs);
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
        private void PrintData(List<LiveStockPrice> data, Dictionary<string, Worksheet> targetWSs)
        {
            if (targetWSs.ContainsKey(""))
            {
                // = 1 ws
                Worksheet targetWS = targetWSs[""];
                string sheetName = targetWS.Name;
                Range usedRange = targetWS.UsedRange;
                int rows = usedRange.Rows.Count;
                if (rows > 1)
                {
                    //old - need to seek for row
                    foreach (var dataRow in data)
                    {
                        bool found = false;
                        for (int row = 2; row <= usedRange.Rows.Count; row++)
                        {
                            if (targetWS.Cells[row, 1].Value == dataRow.Code)
                            {
                                found = true;
                                for (int col = 2; col <= usedRange.Columns.Count; col++)
                                {
                                    if (targetWS.Cells[1, col].Value != null)
                                    {
                                        targetWS.Cells[row, col] = dataRow.GetType().GetProperty(targetWS.Cells[1, col].Value).GetValue(dataRow, null);
                                    }
                                }
                            }
                        }
                        if (!found)
                        {
                            usedRange = targetWS.UsedRange;
                            targetWS.Cells[usedRange.Rows.Count + 1, 1] = dataRow.Code;
                            for (int col = 2; col <= usedRange.Columns.Count; col++)
                            {
                                if (targetWS.Cells[1, col].Value != null)
                                {
                                    targetWS.Cells[usedRange.Rows.Count + 1, col] = dataRow.GetType().GetProperty(targetWS.Cells[1, col].Value).GetValue(dataRow, null);
                                }
                            }
                        }
                    }
                }
                else
                {
                    //new
                    var props = Filters.Where(x => x.Item2 == true).Select(x => x.Item1).ToList();
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
                    Range tableRange = targetWS.Range[targetWS.Cells[1, 1], targetWS.Cells[table.GetLength(0), table.GetLength(1)]];
                    tableRange.Value = table;

                    if (Smart)
                    {
                        MakeTable(tableRange, targetWS, Name, 1);
                    }
                }
            }
            else
            {
                // > 1 ws
                foreach (var item in targetWSs)
                {
                    string ticker = item.Key;
                    Worksheet targetWS = item.Value;
                    var dataRow = data.Find(x => x.Code == ticker);
                    var props = Filters.Where(x => x.Item2 == true).Select(x => x.Item1).ToList();
                    string sheetName = targetWS.Name;
                    Range usedRange = targetWS.UsedRange;
                    int rows = usedRange.Rows.Count;
                    if (rows > 1)
                    {
                        //old - need to move second row down
                        Range secondRow = targetWS.Cells[2, 1].EntireRow;
                        secondRow.Insert(XlInsertShiftDirection.xlShiftDown);
                        object[] tableRow = new object[props.Count];
                        foreach (var prop in props)
                        {
                            tableRow[props.IndexOf(prop)] = dataRow.GetType().GetProperty(prop).GetValue(dataRow, null);
                        }
                        Range tableRowRange = targetWS.Range[targetWS.Cells[2, 1], targetWS.Cells[2, tableRow.Length]];
                        tableRowRange.Value = tableRow;
                    }
                    else
                    {
                        //new
                        object[,] table = new object[2, props.Count];
                        foreach (var prop in props)
                        {
                            table[0, props.IndexOf(prop)] = prop;
                            table[1, props.IndexOf(prop)] = dataRow.GetType().GetProperty(prop).GetValue(dataRow, null);
                        }
                        Range tableRange = targetWS.Range[targetWS.Cells[1, 1], targetWS.Cells[table.GetLength(0), table.GetLength(1)]];
                        tableRange.Value = table;

                        if (Smart)
                        {
                            MakeTable(tableRange, targetWS, Name, 1);
                        }
                    }
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
