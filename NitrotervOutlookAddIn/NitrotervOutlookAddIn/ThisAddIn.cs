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

        public static string path = @"D:\\NitrotervOutlook";

        public static string data_file = @"path.ini";
        public static string projectname_file = "projektszamok.ini";

        static string default_network_path = "D:\\";
        static string default_local_path = "D:\\local_puffer";
        static string default_server_path = "\\\\Nitroterv02server\\Tervezési projektek\\";


        public static string network_path = default_network_path;
        public static string local_path = default_local_path;
        public static string server_path = default_server_path;

        public List<String> yearList;
        public SortedDictionary<String, List<String>> projectNumberList; 

        public string getDefaultServerPath()
        {
            return default_server_path;
        }

        public string getServerPath()
        {
            return server_path;
        }

        public void setServerPath(string value)
        {
            server_path = value;
        }
        public void setPath(string value)
        {
            path = value;
        }

        public string getPath()
        {
            return path;
        }
        public string getProjectnameFile()
        {
            return projectname_file;
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
            string[] lines = { network_path, local_path, server_path };

            if (File.Exists(path + "\\" + data_file))
            {
                File.Delete(path + "\\" + data_file);
            }

            FileStream fs = new FileStream(path + "\\" + data_file, FileMode.Create);

            fs.Close();


            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path + "\\" + data_file))
            {
                foreach (string line in lines)
                {

                    file.WriteLine(line);

                }
            }

            File.SetAttributes(path + "\\" + data_file, FileAttributes.Hidden);
        }

        public void loadFolderNames()
        {
            if (Directory.Exists(server_path))
            {
                yearList = new List<string>();
                projectNumberList = new SortedDictionary<string, List<string>>();

                List<String> folders = new List<string>(Directory.GetDirectories(server_path).Select(d => new DirectoryInfo(d).Name));

                for (int i = 0; i < folders.Count; i++)
                {
                    if (folders[i].Substring(0, 2) == "20")
                    {
                        yearList.Add(folders[i]);
                    }
                }

                foreach (string year in yearList)
                {
                    List<String> projectNumbers = new List<String>(Directory.GetDirectories(server_path + "\\" + year).Select(d => new DirectoryInfo(d).Name));                    

                    for (int i = 0; i < projectNumbers.Count; i++)
                    {
                        projectNumbers[i] = projectNumbers[i].Substring(0, 5);
                    }

                    projectNumberList[year] = projectNumbers;
                }

                //kiiras fajlba
                if (File.Exists(path + "\\" + projectname_file))
                {
                    File.Delete(path + "\\" + projectname_file);
                }

                FileStream fs = new FileStream(path + "\\" + projectname_file, FileMode.Create);

                fs.Close();

                using (StreamWriter file = new StreamWriter(path + "\\" + projectname_file))
                {
                    foreach (string year in yearList)
                    {
                        String line = year;

                        foreach (String projectNumber in projectNumberList[year])
                        {
                            line += " " + projectNumber;
                        }

                        file.WriteLine(line);

                    }
                }

                File.SetAttributes(path + "\\" + projectname_file, FileAttributes.Hidden);
            }
            else if (File.Exists(path + "\\" + projectname_file))
            {
                yearList = new List<string>();
                projectNumberList = new SortedDictionary<string, List<string>>();

                StreamReader file = new StreamReader(path + "\\" + projectname_file);
                int counter = 0;
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    List<String> datas = new List<string>(line.Split(' '));

                    String year = datas[0];

                    yearList.Add(year);
                    datas.Remove(year);

                    projectNumberList[year] = datas;
                }
                file.Close();
            }
            else
            {
                throw new Exception("A szerver és a projektszámokat tartalmazó fájl nem elérhető!");
            }
        }

        public void writeConfigFile()
        {
            //create datafile ini
            if (!File.Exists(path + "\\" + data_file))
            {
                FileStream fs = new FileStream(path + "\\" + data_file, FileMode.Create);

                fs.Close();

                string[] lines = { network_path, local_path, server_path };

                using (StreamWriter file = new StreamWriter(path + "\\" + data_file))
                {
                    foreach (string line in lines)
                    {

                        file.WriteLine(line);

                    }
                }

                File.SetAttributes(path + "\\" + data_file, FileAttributes.Hidden);

            }
            else
            {
                string[] lines = new string[3];
                lines = File.ReadAllLines(path + "\\" + data_file);
                network_path = lines[0];
                local_path = lines[1];
                server_path = lines[2];
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            currentExplorer = this.Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook
                .ExplorerEvents_10_SelectionChangeEventHandler
                (CurrentExplorer_Event);

            try
            {
                //create path
                if (!Directory.Exists(path))
                {
                    var di = Directory.CreateDirectory(path);
                    di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                }


                writeConfigFile();

                loadFolderNames();

            }
            catch (Exception exeption)
            {
                MessageBox.Show("Hiba az adatok olvasása során, kérem ellenőrizze a megadott mappákat.", "Sikertelen adat olvasás.", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        public void saveMailItem(string projectnumber)
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

                    mailItem.SaveAs(nameBuilder(projectnumber, mailItem.Subject, dateBuilder(mailItem.SentOn), dateBuilder(DateTime.Now)));
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

                    mailItem.SaveAs(nameBuilder(projectnumber, mailItem.Subject, dateBuilder(mailItem.SentOn), dateBuilder(DateTime.Now)));
                    mailItem.MarkAsTask(Outlook.OlMarkInterval.olMarkNoDate);

                    MessageBox.Show("Sikertelen hálózati mentés", "Sikertelen iktatásra küldés!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show("Sikeres lokális mentés", "Sikeresen elküldve iktatásra!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                (mailItem as Microsoft.Office.Interop.Outlook._MailItem).Close(Microsoft.Office.Interop.Outlook.OlInspectorClose.olSave);

            }
            catch (UnauthorizedAccessException ex)
            {
                //
                MessageBox.Show("UnauthorizedAccessException\n", "Sikertelen iktatásra küldés!\n", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DirectoryNotFoundException ex)
            {
                //
                MessageBox.Show("DirectoryNotFoundException\n", "Sikertelen iktatásra küldés!\n", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            catch (Exception ex)
            {
                //
                MessageBox.Show("Exception\n" + ex.ToString(), "Sikertelen iktatásra küldés!\n", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private string nameBuilder(string projectnumber, string subject, string sent_date, string date)
        {
            try
            {
                int counter = 1;

                if (subject != null)
                {
                    subject = Regex.Replace(subject, @"\s+", "_");
                    subject = Regex.Replace(subject, "[^\\w\\d]", "_");
                }
                else
                {
                    subject = "";
                }

                
                string name = "[" + projectnumber + "]" + "[" + sent_date + "]" + "[" + date + "]" + "[" + counter + "]" + subject;

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
                    
                    name += "[" + projectnumber + "]" + "[" + sent_date + "]" + "[" + date + "]" + "[" + counter + "]" + subject;

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
