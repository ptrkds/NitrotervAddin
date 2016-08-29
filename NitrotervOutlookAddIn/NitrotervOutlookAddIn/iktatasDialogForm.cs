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
        string project_file;
        public iktatasDialogForm()
        {
            InitializeComponent();
        }
        
        public void setPath()
        {
            local_path = Globals.ThisAddIn.getLocalPath();
            network_path = Globals.ThisAddIn.getNetworkPath();
            project_file = Globals.ThisAddIn.getProjectnameFile();
            localPathTextBox.Text = local_path;
            networkPathTextBox.Text = network_path;
            projectFileTextBox.Text = project_file;
        }

        private void backButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void resetButton_Click(object sender, EventArgs e)
        {
            local_path = Globals.ThisAddIn.getDefaultLocalPath();
            network_path = Globals.ThisAddIn.getDefaultNetworkPath();
            project_file = Globals.ThisAddIn.getDefaultProjectnameFile();

            localPathTextBox.Text = local_path;
            networkPathTextBox.Text = network_path;
            projectFileTextBox.Text = project_file;

        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.setLocalPath(local_path);
            Globals.ThisAddIn.setNetworkPath(network_path);
            Globals.ThisAddIn.setProjectnameFile(project_file);

            Globals.ThisAddIn.dataFileFunction();

            Globals.Ribbons.IktatasMacro.loadProjectFile();

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
                this.networkPathTextBox.Text = folderBrowserDialog2.SelectedPath;
                network_path = this.networkPathTextBox.Text;
            }
        }

        private void projectFileTextBox_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Text Files (.txt)|*.txt|All Files (*.*)|*.*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                this.projectFileTextBox.Text = openFileDialog1.FileName;
                project_file = this.projectFileTextBox.Text;
            }
        }
    }
}
