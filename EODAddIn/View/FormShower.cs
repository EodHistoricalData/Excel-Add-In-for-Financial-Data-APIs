using EODAddIn.BL.Live;
using EODAddIn.Forms;
using EODAddIn.Program;
using EODAddIn.Utils;

using Microsoft.Office.Core;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace EODAddIn.View.Forms
{
    internal class FormShower
    {
        private static readonly Dictionary<(string, string), Form> _forms = new Dictionary<(string, string), Form>();

        public static void FrmGetBulkShow() => FrmShow(typeof(FrmGetBulk));
        public static void FrmGetBulkEodShow() => FrmShow(typeof(FrmGetBulkEod));
        public static void FrmGetEtfShow() => FrmShow(typeof(FrmGetEtf));
        public static void FrmGetFundamentalShow() => FrmShow(typeof(FrmGetFundamental));
        public static void FrmGetHistoricalShow() => FrmShow(typeof(FrmGetHistorical));
        public static void FrmGetIntradayHistoricalDataShow() => FrmShow(typeof(FrmGetIntradayHistoricalData));
        public static void FrmGetLiveShow() => FrmShow(typeof(FrmGetLive));
        public static void FrmGetOptionsShow() => FrmShow(typeof(FrmGetOptions));
        public static void FrmGetTechnicalsShow() => FrmShow(typeof(FrmGetTechnicals));
        public static void FrmLiveFiltersShow() => FrmShow(typeof(FrmLiveFilters));
        public static void FrmScreenerShow() => FrmShow(typeof(FrmScreener));
        public static void FrmScreenerHistoricalShow() => FrmShow(typeof(FrmScreenerHistorical));
        public static void FrmScreenerIntradayShow() => FrmShow(typeof(FrmScreenerIntraday));

        public static LiveDownloaderDispatcher LiveDownloaderDispatcherShow(Dictionary<LiveDownloader, CustomXMLPart> downloaders, Dictionary<LiveDownloader, CancellationTokenSource> cancellationTokenSources)
        {
            if (ShowActiveForm()) throw new DocumentAlreadyLoadedException();

            LiveDownloaderDispatcher frm;

            if (downloaders.Count == 0)
            {
                frm = new LiveDownloaderDispatcher();
            }
            else
            {
                frm = new LiveDownloaderDispatcher(downloaders, cancellationTokenSources);
            }

            FrmShow(frm);
            return frm;
        }

        public static bool ShowActiveForm()
        {
            string appHwnd = Globals.ThisAddIn.Application.Hwnd.ToString();
            var formsApp = _forms.Where(x => x.Key.Item1 == appHwnd);
            if (formsApp.Count() > 0)
            {
                formsApp.First().Value.Activate();
                return true;
            }
            return false;
        }

        private static void FrmShow(Type formType)
        {
            try
            {
                var key = (Globals.ThisAddIn.Application.Hwnd.ToString(), formType.Name);
                

                Form form;
                if (_forms.Count > 0)
                {
                    form = _forms.First().Value;
                    form.Activate();
                    return;
                }
                if (_forms.ContainsKey(key))
                {
                    form = _forms[key];
                    form.Activate();
                }
                else
                {
                    form = (Form)Activator.CreateInstance(formType);
                    _forms.Add(key, form);
                    form.FormClosed += Form_FormClosed;
                    form.Show(new WinHwnd());
                }
            }
            catch (ViewException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                new ErrorReport(ex).ShowAndSend();
            }
        }

        public static void FrmShow(Form form)
        {
            try
            {
                var key = (Globals.ThisAddIn.Application.Hwnd.ToString(), form.GetType().Name);
               // string key = Globals.ThisAddIn.Application.Hwnd.ToString() + form.GetType().Name;
                if (_forms.Count > 0)
                {
                    form = _forms.First().Value;
                    form.Activate();
                    return;
                }
                if (_forms.ContainsKey(key))
                {
                    form = _forms[key];
                    form.Activate();
                }
                else
                {
                    form.FormClosed += Form_FormClosed;
                    form.Show(new WinHwnd());
                    _forms.Add(key, form);
                }
            }
            catch (ViewException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                new ErrorReport(ex).ShowAndSend();
            }
        }

        private static void Form_FormClosed(object sender, FormClosedEventArgs e)
        {
            var key = (Globals.ThisAddIn.Application.Hwnd.ToString(), sender.GetType().Name);
            _forms.Remove(key);
        }

    }
}
