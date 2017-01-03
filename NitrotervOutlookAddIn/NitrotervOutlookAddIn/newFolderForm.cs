using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NitrotervOutlookAddIn
{
    public partial class newFolderForm : Form
    {
        private string yearFolder;
        public newFolderForm( string _yearFolder )
        {
            yearFolder = _yearFolder;
            InitializeComponent();
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (Directory.Exists(Globals.ThisAddIn.getServerPath()))
                {
                    String folder = Globals.ThisAddIn.getServerPath() + "\\" + yearFolder;
                    if (!Regex.IsMatch(folderNameTextBox.Text, @"^\d+$"))
                    {
                        MessageBox.Show("Csak számot adjon meg projektszámnak!", "Hiba a mappa létrehozása során", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                    else if (folderNameTextBox.Text == "")
                    {
                        MessageBox.Show("Adjon meg projektszámot", "Hiba a mappa létrehozása során", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        folder += "\\" + folderNameTextBox.Text;
                        if (!Directory.Exists(folder))
                        {
                            Directory.CreateDirectory(folder);
                            if (folderNameTextBox.Text.Substring((folderNameTextBox.Text.Length - 3), 3) != "001")
                            {
                                folder += "\\03 Adminisztráció";
                                if (!Directory.Exists(folder))
                                {
                                    Directory.CreateDirectory(folder);
                                }

                                string incoming = folder + "\\Beérkező levelek ";
                                if (!Directory.Exists(incoming))
                                {
                                    Directory.CreateDirectory(incoming);
                                }
                                string outgoing = folder + "\\Kimenő levelek ";
                                if (!Directory.Exists(outgoing))
                                {
                                    Directory.CreateDirectory(outgoing);
                                }
                            }
                            MessageBox.Show("Mappa sikeresen létrehozva", "Mappa sikeresen létrehozva", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        }
                        else
                        {
                            MessageBox.Show("Már létezik a mappa", "Hiba a mappa létrehozása során", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }


                    }
                }
                
                DialogResult = DialogResult.OK;

            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Hiba a mappa létrehozása során", "Hiba a mappa létrehozása során", MessageBoxButtons.OK, MessageBoxIcon.Error);

                DialogResult = DialogResult.Abort;
            }



        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}
