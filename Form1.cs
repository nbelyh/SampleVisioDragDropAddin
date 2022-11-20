using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace DragDropAddin
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void listView1_ItemDrag(object sender, ItemDragEventArgs e)
        {
            var item = (ListViewItem)e.Item;

            var app = Globals.ThisAddIn.Application;

            var myStencil = app.Documents.OpenEx("BASIC_M.VSS",
                (short)Visio.VisOpenSaveArgs.visOpenDocked | (short)Visio.VisOpenSaveArgs.visOpenRO);

            var masterToDrag = myStencil.Masters[item.Text];
            var data = new DataObject(masterToDrag);
            DoDragDrop(data, DragDropEffects.Copy);
        }
    }
}
