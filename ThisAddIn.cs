using System;
using Microsoft.Office.Interop.Visio;
using System.Windows.Forms;

namespace DragDropAddin
{
    public partial class ThisAddIn
    {
        Form1 _form1 = new Form1();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.Documents.Add("");
            _form1.Show(NativeWindow.FromHandle((IntPtr)Application.Window.WindowHandle32));
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _form1.Close();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
