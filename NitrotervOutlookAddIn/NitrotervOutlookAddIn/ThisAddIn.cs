using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace NitrotervOutlookAddIn
{
    public partial class ThisAddIn
    {
        private Outlook.Explorer currentExplorer = null;
        Outlook.MailItem mailItem;

        public static string data_file = @"D:\\path.txt";

        static string default_network_path = "\\\\Nitroterv02\\e\\Tervezesi projektek\\2016\\16017 NZrt Pétisó üzem bővítés\\03 Adminisztráció\\_iktatásra";
        static string default_local_path = "D:\\local_puffer";

        private static string default_projectname_file =
            "D:\\projektnyilvántartás.txt";


        public static string network_path = default_network_path;
        public static string local_path = default_local_path;
        public static string projectname_file = default_projectname_file;


        public string getProjectnameFile()
        {
            return projectname_file;
        }
        public void setProjectnameFile(string value)
        {
            projectname_file = value;
        }

        public string getDefaultProjectnameFile()
        {
            return default_projectname_file;
        }

        public string getDefaultLocalPath()
        {
            return default_local_path;
        }
        public string getDefaultNetworkPath()
        {
            return default_network_path;
        }

        public string getLocalPath()
        {
            return local_path;
        }
        public void setLocalPath(string _local_path)
        {
            local_path = _local_path;
        }
        public string getNetworkPath()
        {
            return network_path;
        }
        public void setNetworkPath(string _network_path)
        {
            network_path = _network_path;
        }

        public void dataFileFunction()
        {
            string[] lines = { network_path, local_path, projectname_file };

            if (File.Exists(data_file))
            {
                File.Delete(data_file);
            }

            FileStream fs = new FileStream(data_file, FileMode.Create);

            fs.Close();


            using (System.IO.StreamWriter file = new System.IO.StreamWriter(data_file))
            {
                foreach (string line in lines)
                {

                    file.WriteLine(line);

                }
            }

            File.SetAttributes(data_file, FileAttributes.Hidden);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            currentExplorer = this.Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook
                .ExplorerEvents_10_SelectionChangeEventHandler
                (CurrentExplorer_Event);

            try
            {
                if (!File.Exists(data_file))
                {
                    FileStream fs = new FileStream(data_file, FileMode.Create);

                    fs.Close();

                    string[] lines = { network_path, local_path, projectname_file };

                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(data_file))
                    {
                        foreach (string line in lines)
                        {

                            file.WriteLine(line);
                            
                        }
                    }

                    File.SetAttributes(data_file, FileAttributes.Hidden);

                }
                else
                {
                    string[] lines = new string[3];
                    lines = File.ReadAllLines(data_file);
                    network_path = lines[0];
                    local_path = lines[1];
                    projectname_file = lines[2];
                }
            }
            catch (Exception exeption)
            {
                MessageBox.Show(exeption.ToString(), "Sikertelen data_file olvasás.", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            //local check
            LocalToNetwork();

        }

        public bool IsDirectoryEmpty(string path)
        {
            return !Directory.EnumerateFileSystemEntries(path).Any();
        }


        public void LocalToNetwork()
        {
            try
            {
                if (System.IO.Directory.Exists(network_path))
                {

                    if (!Directory.Exists(local_path))
                    {
                        DirectoryInfo di = Directory.CreateDirectory(local_path);
                        di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                    }

                    if (!IsDirectoryEmpty(local_path))
                    {
                        List<String> local_files = Directory.GetFiles(local_path, "*.*", SearchOption.AllDirectories).ToList();

                        foreach (string file in local_files)
                        {
                            FileInfo mFile = new FileInfo(file);

                            string path = network_path + "\\" + Path.GetFileNameWithoutExtension(mFile.Name);

                            // to remove name collusion
                            if (new FileInfo(network_path + "\\" + mFile.Name).Exists == false)

                                mFile.MoveTo(network_path + "\\" + mFile.Name);

                            else
                            {
                                int counter = 1;
                                string saveAs = path + "_" + counter + ".msg";

                                while (new FileInfo(saveAs).Exists == true)
                                {
                                    counter ++;
                                    saveAs = path + "_" + counter + ".msg";
                                }

                                mFile.MoveTo(saveAs);

                            }
                        }

                        MessageBox.Show("Sikeres továbbitás a hálózati iktatás mappába!", "Sikeres művelet", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
                else
                {
                    MessageBox.Show("Nincs hálózati kapcsolat!\n", "Hálózati hiba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception!\n" + ex.ToString(), "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
        private void CurrentExplorer_Event()
        {
            Outlook.MAPIFolder selectedFolder =
                this.Application.ActiveExplorer().CurrentFolder;

            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];
                    if (selObject is Outlook.MailItem)
                    {
                        mailItem = (selObject as Outlook.MailItem);
                    }
                    else
                    {
                        mailItem = null;
                    }
                }
            }
            catch (Exception ex)
            {
                //expMessage = ex.Message;
                MessageBox.Show("Exception\n" + ex.ToString(), "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        public void saveMailItem(string project)
        {
            string[] dirs = { "" };

            if (mailItem == null)
                return;

            try
            {
                if (System.IO.Directory.Exists(network_path))
                {
                    //finds the project dir where path ends with "*project"
                    //dirs = Directory.GetDirectories(network_path, "*" + project, System.IO.SearchOption.AllDirectories);
                    //Note: 
                    // .../dir1/asdproject and .../dir1/project are also found!!

                    mailItem.SaveAs(nameBuilder(project, mailItem.Subject, dateBuilder(mailItem.SentOn)));
                    //mailItem.SaveAs(dirs[0] + "\\" + nameBuilder(project, mailItem.Subject) + ".msg");
                    //mailItem.Categories = "Iktatva";
                    try
                    {
                        var customCat = "Iktatásra küldve";
                        if (Application.Session.Categories[customCat] == null)
                            Application.Session.Categories.Add(customCat, Outlook.OlCategoryColor.olCategoryColorDarkRed);

                        mailItem.Categories = customCat;
                        //mailItem.MarkAsTask(Outlook.OlMarkInterval.olMarkNoDate);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Sikertelen kategorizálás", "Sikertelen kategorizálás!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    MessageBox.Show("Sikeres hálózati mentés", "Sikeresen elküldve iktatásra!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    //save file to local buffer


                    if (!Directory.Exists(local_path))
                    {
                        DirectoryInfo di = Directory.CreateDirectory(local_path);
                        di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                    }

                    mailItem.SaveAs(nameBuilder(project, mailItem.Subject, dateBuilder(mailItem.SentOn)));
                    mailItem.MarkAsTask(Outlook.OlMarkInterval.olMarkNoDate);

                    MessageBox.Show("Sikertelen hálózati mentés", "Sikertelen iktatásra küldés!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show("Sikeres lokális mentés", "Sikeresen elküldve iktatásra!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                //
                MessageBox.Show("UnauthorizedAccessException\n" + ex.ToString(), "Sikertelen iktatásra küldés!\n", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DirectoryNotFoundException ex)
            {
                //
                MessageBox.Show("DirectoryNotFoundException\n" + ex.ToString(), "Sikertelen iktatásra küldés!\n", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            catch (Exception ex)
            {
                //
                MessageBox.Show("Exception\n" + ex.ToString(), "Sikertelen iktatásra küldés!\n", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private string nameBuilder(string project, string subject, string date)
        {
            try
            {
                int counter = 1;

                if (subject != null)
                {
                    subject = Regex.Replace(subject, @"\s+", "_");
                    subject = Regex.Replace(subject, "[^\\w\\d]", "");
                }
                else
                {
                    subject = "";
                }

                string name = "[" + project + "]" + "[" + date + "]" + "[" + counter + "]" + subject;

                string path;
                if (System.IO.Directory.Exists(network_path))
                {
                    path = network_path + "\\" + name;
                }
                else
                {
                    path = local_path + "\\" + name;
                }
                if (path.Length > 245)
                    path = path.Remove(245);

                path += ".msg";

                while (File.Exists(path))
                {
                    counter++;
                    name = "[" + project + "]" + "[" + date + "]" + "[" + counter + "]" + subject;

                    path = "";

                    if (System.IO.Directory.Exists(network_path))
                    {
                        path = network_path + "\\" + name;
                    }
                    else
                    {
                        path = local_path + "\\" + name;
                    }
                    if (path.Length > 245)
                        path = path.Remove(245);

                    path += ".msg";
                }

                return path;
            }
            catch (Exception)
            {
                throw;
            }  
        }

        private string dateBuilder(DateTime date)
        {
            try
            {
                return String.Format("{0:yyMMdd}", date);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /*private string pathBuilder(string project)
        {
            try
            {
                

            }
            catch (Exception)
            {
                throw;
            }
        }*/

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
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
