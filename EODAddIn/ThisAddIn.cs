using EODAddIn.BL;
using EODAddIn.Program;
using System;
using System.IO;

namespace EODAddIn
{
    public partial class ThisAddIn
    {
        private UDF utilities;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Program.Program.Run();
            EODAddIn.Panels.PanelInfo panel = new Panels.PanelInfo(); // add to settings save 
            panel.ShowPanel();
            Settings.SettingsFields.IsInfoShowed = true;
            string filename = Path.Combine(Program.Program.UserFolder, "EODAddIn.xla");
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException("Не найден файл с функциями EODAddIn.xla");
            }
            try
            {
                Globals.ThisAddIn.Application.Workbooks.Open(filename);
            }
            catch
            {
                throw;
            }

            try
            {
                string oldSettingsFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), Program.Program.CompanyName, "EOD Excel Plagin", "settings.xml");
                string newSettingsFile = Path.Combine(Program.Program.UserFolder, "settings.xml");

                if (File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), Program.Program.CompanyName, "EOD Excel Plagin", "settings.xml")))
                {
                    File.Copy(oldSettingsFile, newSettingsFile);
                    File.Delete(oldSettingsFile);
                    Directory.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), Program.Program.CompanyName, "EOD Excel Plagin"));
                }
            }
            catch (System.Exception)
            {

            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /// <summary>
        /// Переопределение утилит для взаимодействия функций из VBA
        /// </summary>
        /// <returns></returns>
        protected override object RequestComAddInAutomationService()
        {
            if (utilities == null) utilities = new UDF();
            return utilities;
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
