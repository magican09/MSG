using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MSGAddIn
{
    public partial class SidePane : UserControl
    {
        public SidePane()
        {
          //  InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", "-p");
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // SidePane
            // 
            this.Name = "SidePane";
            this.Size = new System.Drawing.Size(470, 348);
            this.ResumeLayout(false);

        }
    }
}
