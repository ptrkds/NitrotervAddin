using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace NitrotervOutlookAddIn
{
    partial class IktatasMacro : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public IktatasMacro()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.iktatasTab = this.Factory.CreateRibbonTab();
            this.iktatasGroup = this.Factory.CreateRibbonGroup();
            this.yearDropDown = this.Factory.CreateRibbonDropDown();
            this.projektekDropDown = this.Factory.CreateRibbonDropDown();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.iktatasButton = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.localCheckButton = this.Factory.CreateRibbonButton();
            this.networkFolderButton = this.Factory.CreateRibbonButton();
            this.iktatasTab.SuspendLayout();
            this.iktatasGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // iktatasTab
            // 
            this.iktatasTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.iktatasTab.Groups.Add(this.iktatasGroup);
            this.iktatasTab.Label = "Küldés iktatásra";
            this.iktatasTab.Name = "iktatasTab";
            // 
            // iktatasGroup
            // 
            this.iktatasGroup.DialogLauncher = ribbonDialogLauncherImpl1;
            this.iktatasGroup.Items.Add(this.yearDropDown);
            this.iktatasGroup.Items.Add(this.projektekDropDown);
            this.iktatasGroup.Items.Add(this.separator1);
            this.iktatasGroup.Items.Add(this.iktatasButton);
            this.iktatasGroup.Items.Add(this.separator2);
            this.iktatasGroup.Items.Add(this.localCheckButton);
            this.iktatasGroup.Items.Add(this.networkFolderButton);
            this.iktatasGroup.Label = "Iktatás";
            this.iktatasGroup.Name = "iktatasGroup";
            // 
            // yearDropDown
            // 
            this.yearDropDown.Label = "Év               ";
            this.yearDropDown.Name = "yearDropDown";
            this.yearDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.yearDropDown_SelectionChanged);
            // 
            // projektekDropDown
            // 
            this.projektekDropDown.Label = "Projektszám";
            this.projektekDropDown.Name = "projektekDropDown";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // iktatasButton
            // 
            this.iktatasButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.iktatasButton.Image = global::NitrotervOutlookAddIn.Properties.Resources.file_save;
            this.iktatasButton.Label = "Küldés iktatásra";
            this.iktatasButton.Name = "iktatasButton";
            this.iktatasButton.ShowImage = true;
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // localCheckButton
            // 
            this.localCheckButton.Image = global::NitrotervOutlookAddIn.Properties.Resources.folder_check;
            this.localCheckButton.Label = "Lokális Mappa Ellenőrzése";
            this.localCheckButton.Name = "localCheckButton";
            this.localCheckButton.ShowImage = true;
            // 
            // networkFolderButton
            // 
            this.networkFolderButton.Image = global::NitrotervOutlookAddIn.Properties.Resources.folder_check;
            this.networkFolderButton.Label = "Hálózati Mappa Megnyitása";
            this.networkFolderButton.Name = "networkFolderButton";
            this.networkFolderButton.ShowImage = true;
            // 
            // IktatasMacro
            // 
            this.Name = "IktatasMacro";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.iktatasTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.IktatasMacro_Load);
            this.iktatasTab.ResumeLayout(false);
            this.iktatasTab.PerformLayout();
            this.iktatasGroup.ResumeLayout(false);
            this.iktatasGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab iktatasTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup iktatasGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown projektekDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton iktatasButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton localCheckButton;

        private void iktatasButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string selectedProject;

                if (projektekDropDown.SelectedItem != null)
                {
                    selectedProject = projektekDropDown.SelectedItem.Label;
                }
                else
                {
                    selectedProject = yearDropDown.SelectedItem.Label;
                }

                Globals.ThisAddIn.saveMailItem(selectedProject);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Nem létezik a projektnyilvántartási file.", "File not found", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }


        private void localCheckButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.LocalToNetwork();
        }

        private void networkFolderButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (System.IO.Directory.Exists(Globals.ThisAddIn.getNetworkPath()))
            {
                Process.Start(Globals.ThisAddIn.getNetworkPath());
            }
            else
            {
                Process.Start(Globals.ThisAddIn.getLocalPath());
            }

        }

        private void iktatasGroup_DialogLauncherClick(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            iktatasDialogForm iktatas_dialog = new iktatasDialogForm();

            iktatas_dialog.setPath();
            iktatas_dialog.Show();
        }

        private void yearDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            projektekDropDown.Items.Clear();
            foreach (String projectNumber in Globals.ThisAddIn.projectNumberList[yearDropDown.SelectedItem.Label])
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = projectNumber;
                projektekDropDown.Items.Add(item);
            }
        }

        internal RibbonButton networkFolderButton;
        internal RibbonSeparator separator1;
        internal RibbonSeparator separator2;
        internal RibbonDropDown yearDropDown;
    }

    partial class ThisRibbonCollection
    {
        internal IktatasMacro IktatasMacro
        {
            get { return this.GetRibbon<IktatasMacro>(); }
        }
    }
}
