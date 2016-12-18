using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Windows.Forms;

namespace NitrotervOutlookAddIn
{
    public partial class IktatasMacro
    {
        public void loadProjectFile()
        {
            projektekDropDown.Items.Clear();

            //combobox from file
            int counter = 0;
            string line;
            try
            {
                //StreamReader file = new StreamReader(Globals.ThisAddIn.getPath() + "\\projektnyilvántartás.txt");
                StreamReader file = new StreamReader(Globals.ThisAddIn.getProjectnameFile());

                while ((line = file.ReadLine()) != null)
                {
                    RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    item.Label = line;
                    projektekDropDown.Items.Add(item);
                    counter++;
                }
                file.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Nem létezik a projektnyilvántartási file.", "File not found", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void IktatasMacro_Load(object sender, RibbonUIEventArgs e)
        {


            loadProjectFile();


            //click event
            this.iktatasButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.iktatasButton_Click);
            this.localCheckButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.localCheckButton_Click);
            this.iktatasGroup.DialogLauncherClick += new RibbonControlEventHandler(this.iktatasGroup_DialogLauncherClick);
            this.networkFolderButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.networkFolderButton_Click);

        }


    }
}
