using Microsoft.Office.Tools;

using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace EODAddIn.Panels
{
    /// <summary>
    /// Базовый класс панелей Excel
    /// </summary>
    public partial class PanelExcel : UserControl
    {
        public Excel.Workbook Workbook;

        /// <summary>
        /// Видимость панели
        /// </summary>
        public bool VisiblePanel
        {
            get
            {
                if (CustomTaskPanel != null)
                {
                    return CustomTaskPanel.Visible;
                }
                return false;
            }
            set
            {
                CustomTaskPanel.Visible = value;
            }
        }

        /// <summary>
        /// Заголовок
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Панель word 
        /// </summary>
        public CustomTaskPane CustomTaskPanel;

        public PanelExcel()
        {

        }

        /// <summary>
        /// Конструктор панели. 
        /// </summary>
        /// <param name="title">Подпись панели</param>
        public PanelExcel(string title)
        {
            Workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Title = title;
            CustomTaskPanel = Globals.ThisAddIn.CustomTaskPanes.Add(this, Title);
        }

        /// <summary>
        /// Отобразить панель
        /// </summary>
        public void ShowPanel()
        {
            CustomTaskPanel.Visible = true;
        }

        /// <summary>
        /// Скрыть панель
        /// </summary>
        public void HidePanel()
        {
            CustomTaskPanel.Visible = false;
        }

    }
}
