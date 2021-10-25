using System;
using System.Drawing;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace ProXL.Forms
{
    public partial class RefEdit : UserControl
    {
        private readonly Excel.Application Application = Globals.ThisAddIn?.Application;
        public struct ControlState
        {
            public Size ClientSize;
            public bool IsMinimized;
            public int ControlLeft;
            public int ControlTop;
            public int ControlWidth;
            public AnchorStyles ControlAnchor;
        }

        public new string Text
        {
            get
            {
                return txtRef.Text;
            }
            set
            {
                txtRef.Text = value;
            }
        }
        public ControlState State;

        public new event Action TextChanged;

        public RefEdit()
        {
            InitializeComponent();
        }

        private void TxtRef_Enter(object sender, EventArgs e)
        {
            Application.SheetSelectionChange += Application_SheetSelectionChange;
        }
        private void TxtRef_Leave(object sender, EventArgs e)
        {
            Application.SheetSelectionChange -= Application_SheetSelectionChange;
        }

        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            txtRef.Text = Target.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
        }

        private void ResizeForm(object sender, EventArgs e)
        {
            if (!State.IsMinimized) SaveState();
            State.IsMinimized = !State.IsMinimized;

            foreach (Control control in ParentForm.Controls)
            {
                control.Visible = !State.IsMinimized;
            }

            this.Visible = true;

            if (State.IsMinimized)
            {
                ParentForm.ClientSize = new Size(this.Width, this.Height);
                this.Left = 0;
                this.Top = 0;
                this.Width = ParentForm.ClientSize.Width;
                this.Anchor = AnchorStyles.Left;

            }
            else
            {
                ParentForm.ClientSize = State.ClientSize;
                this.Left = State.ControlLeft;
                this.Top = State.ControlTop;
                this.Width = State.ControlWidth;
                this.Anchor = State.ControlAnchor;
            }
            btnUp.Visible = !State.IsMinimized;
            btnDown.Visible = State.IsMinimized;
            Focus();
        }

        private void SaveState()
        {
            State.ClientSize = ParentForm.ClientSize;
            State.ControlLeft = this.Left;
            State.ControlTop = this.Top;
            State.ControlWidth = this.Width;
            State.ControlAnchor = this.Anchor;
        }

        public new void Focus()
        {
            txtRef.Focus();
        }

        private void txtRef_TextChanged(object sender, EventArgs e)
        {
            TextChanged?.Invoke();
        }
    }
}
