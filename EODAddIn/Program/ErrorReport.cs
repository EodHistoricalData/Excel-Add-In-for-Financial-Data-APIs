using EODAddIn.Utils;

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EODAddIn.Program
{
    /// <summary>
    /// Class for displaying errors to the user
    /// </summary>
    public class ErrorReport
    {
        private readonly Exception Exception;
        private Dictionary<string, string> Pairs = new Dictionary<string, string>();
        private const string Url = "https://micro-solution.ru/api/programs/error-report.php";

        /// <summary>
        /// New error constructor
        /// </summary>
        /// <param name="ex"></param>
        public ErrorReport(Exception ex)
        {
            Exception = ex;
        }

        /// <summary>
        /// Constructor for multiple ticker errors
        /// </summary>
        /// <param name="pairs">(ticker, message)</param>
        public ErrorReport(Dictionary<string, string> pairs)
        {
            Pairs = pairs;
        }

        /// <summary>
        /// Message to the user (without sending to the server)
        /// </summary>
        /// <param name="messageAdd"></param>
        public void MessageToUser(string messageAdd = "")
        {
            MessageBox.Show($"Message: {Exception.Message}\n{messageAdd}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Show an error message to the user and send the message to the server
        /// </summary>
        public void ShowAndSend()
        {
            if (Pairs.Count == 0)
            {
                MessageBox.Show($"An error occurred in the program\n" +
                                $"We will receive a report on it and will try to fix the error as soon as possible..\n\n" +
                                $"Error message: {Exception.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Send();
            }
            else
            {
                string tickers = "";
                string message = "";
                foreach (var pair in Pairs)
                {
                    tickers += pair.Key + ", ";
                    message += pair.Key + ": " + pair.Value + ". ";
                }
                tickers = tickers.Substring(0, tickers.Length - 2);
                if (message.Length > 450)
                {
                    message = message.Substring(0, 450);
                    message += "...";
                }
                DialogResult result = MessageBox.Show($"Do you wish to upload the list of errors to a sheet \"Error tickers\"?\n\n" +
                                                    $"Data for certain tickers was not downloaded. Please double-check that there are no misspellings in the ticker name or exchange code and try again. Contact our support team if the error persists.\n\n" +
                                                    $"Failed tickers: {tickers}\n" +
                                                    $"Error message: {message}", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);

                if (result == DialogResult.Yes)
                {
                    var worksheet = ExcelUtils.AddSheet("Error tickers");
                    worksheet.Cells[1, 1].value = "Ticker";
                    worksheet.Cells[1, 2].value = "Error";
                    int row = 2;
                    foreach (var pair in Pairs)
                    {
                        worksheet.Cells[row, 1].value = pair.Key;
                        worksheet.Cells[row, 2].value = pair.Value;
                        row++;
                    }
                }
            }
        }

        /// <summary>
        /// Sending an error message to the server
        /// </summary>
        public void Send()
        {
            Task.Factory.StartNew(() =>
            {
                try
                {
                    Response.POST(Url, $"prog_id=1001&" +
                                        $"prog_ver={Program.Version.Text}&" +
                                        $"comp={Program.UserHash}&" +
                                        $"error={Exception}");
                }
                catch { }
            });
        }
    }
}
