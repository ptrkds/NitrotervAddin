using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NitrotervOutlookAddIn
{
    public partial class iktatasDialogForm : Form
    {
        string local_path;
        string network_path;
        public iktatasDialogForm()
        {
            InitializeComponent();
        }
        
        public void setPath()
        {
            local_path = Globals.ThisAddIn.getLocalPath();
            network_path = Globals.ThisAddIn.getNetworkPath();
            localPathTextBox.Text = local_path;
            networkPathTextBox.Text = network_path;
        }

        private void backButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void resetButton_Click(object sender, EventArgs e)
        {
            local_path = Globals.ThisAddIn.getDefaultLocalPath();
            network_path = Globals.ThisAddIn.getDefaultNetworkPath();

            localPathTextBox.Text = local_path;
            networkPathTextBox.Text = network_path;
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.setLocalPath(local_path);
            Globals.ThisAddIn.setNetworkPath(network_path);

            this.Close();
        }

        private void localPathTextBox_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.localPathTextBox.Text = folderBrowserDialog1.SelectedPath;
                local_path = this.localPathTextBox.Text;
            }
        }

        private void networkPathTextBox_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog2.ShowDialog() == DialogResult.OK)
            {
                this.networkPathTextBox.Text = folderBrowserDialog1.SelectedPath;
                network_path = this.networkPathTextBox.Text;
            }
        }
    }
}
