using EODAddIn.Utils;

using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EODAddIn.Program
{
    /// <summary>
    /// Класс отображения ошибок пользователю
    /// </summary>
    public class ErrorReport
    {
        private readonly Exception Exception;
        private const string Url = "https://micro-solution.ru/api/programs/error-report.php";

        /// <summary>
        /// Конструктор новой ошибки
        /// </summary>
        /// <param name="ex"></param>
        public ErrorReport(Exception ex)
        {
            Exception = ex;
        }

        /// <summary>
        /// Сообщение пользователю (без отправки на сервер)
        /// </summary>
        /// <param name="messageAdd"></param>
        public void MessageToUser(string messageAdd = "")
        {
            MessageBox.Show($"Message: {Exception.Message}\n{messageAdd}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Отображение сообщения об ошибке пользователю и отправка сообщения
        /// </summary>
        public void ShowAndSend()
        {
            MessageBox.Show($"An error occurred in the program\n" +
                            $"We will receive a report on it and will try to fix the error as soon as possible..\n\n" +
                            $"Error message: {Exception.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Send();
        }

        /// <summary>
        /// Отправка сообщения об ошибке на сервер
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
