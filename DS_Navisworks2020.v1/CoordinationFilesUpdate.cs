using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

//User spaces
using DS_Space;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DS_NWClass
{
    [Plugin("DS_CoordinationFilesUpdate_v1.4", "DS", ToolTip = "NWC files assembling to NWD", DisplayName = "DS_CoordinationFilesUpdate_v1.4")]

    public class NWC_Assembly_Plugin : AddInPlugin

    {
        //Get current date and time 
        readonly string CurDate = DateTime.Now.ToString("yyMMdd");
        readonly string CurDateTime = DateTime.Now.ToString("yyMMdd_HHmmss");
        public int FileSize { get; private set; }
        public int FileDate { get; private set; }

        public override int Execute(params string[] parameters)
        {

            FormsInvoke filters_Applying = new FormsInvoke();

            if (filters_Applying.FolderPathNWCpr != null && filters_Applying.FolderPathNWDpr != null)
            {
                filters_Applying.Show();
            }

            return 0;
        }

        public void MainProgram(string FolderPathNWC, string FolderPathNWD, int FileSizeOut, int FileDateOut)
        {
            FileSize = FileSizeOut;
            FileDate = FileDateOut;
            //create NavisworksApplication automation objects   
            Autodesk.Navisworks.Api.Automation.NavisworksApplication automationApplication = null;

            //Intiating main process


            Task task = DirIterateAsync(FolderPathNWC, FolderPathNWD, automationApplication);

        }

        private async Task EndTaskAsync()
        {
            await Task.Run(() =>
            {
                if (File.Exists(LogPath(CurDateTime)) == true)
                {
                    MessageBox.Show("Process has been stoped because errors occured!" + "\n" + "Log saved: " + LogPath(CurDateTime));
                    Clipboard.SetText(LogPath(CurDateTime));
                    return;
                }
                MessageBox.Show("Done!");
            });
        }

        void InitiateFileOperations(string ZeroIndEl, string[] FilesList, string DirPathNWD, string dirName, string FileNameNWD)
        {
            //create NavisworksApplication automation object 
            Autodesk.Navisworks.Api.Automation.NavisworksApplication automationApplication =
               new Autodesk.Navisworks.Api.Automation.NavisworksApplication();

            automationApplication.OpenFile(ZeroIndEl, FilesList);

            string NWDDir = DirPathNWD + "\\" + dirName + "\\";
            if (Directory.Exists(NWDDir) == false)
            {
                Directory.CreateDirectory(NWDDir);
            }

            automationApplication.SaveFile(NWDDir + "\\" + FileNameNWD);

            Archiving(NWDDir, FileNameNWD);
        }

        public async Task DirIterateAsync(string DirPathNWC, string DirPathNWD, Autodesk.Navisworks.Api.Automation.NavisworksApplication automationApplication)
        {

            string[] NewDir = Directory.EnumerateDirectories(DirPathNWC, "*_*_*_*_*", SearchOption.AllDirectories).ToArray();
            int i = 0;
            try
            {
                //Get folders
                foreach (string d in NewDir)
                {
                    i++;
                    // Make a reference to info of a directory.     
                    DirectoryInfo di = new DirectoryInfo(d);

                    // Get a reference to each file in that directory.
                    FileInfo[] fiArr = di.GetFiles();


                    if (fiArr.Length != 0)
                    {
                        string dirName = new DirectoryInfo(d).Name;
                        string FileNameNWD = dirName + "_" + CurDate + ".nwd";

                        string[] FilesList = new string[fiArr.Length];

                        FilesList = GetFilesList(d, fiArr, out string ZeroIndEl, FilesList);

                        if (ZeroIndEl != "")
                        {

                            try
                            {
                                await Task.Run(() => InitiateFileOperations(ZeroIndEl, FilesList, DirPathNWD, dirName, FileNameNWD));

                            }
                            catch (Autodesk.Navisworks.Api.Automation.AutomationException e)
                            {
                                //An error occurred, display it to the user
                                System.Windows.Forms.MessageBox.Show("Error: " + e.Message);
                            }
                            catch (Autodesk.Navisworks.Api.Automation.AutomationDocumentFileException e)
                            {
                                //An error occurred, display it to the user
                                System.Windows.Forms.MessageBox.Show("Error: " + e.Message);
                            }
                            finally
                            {
                                if (automationApplication != null)
                                {
                                    automationApplication.Dispose();
                                    automationApplication = null;
                                }
                            }
                        }


                    }

                }
            }
            catch (Exception ex)
            {
                LogWriter(ex.ToString(), CurDateTime);
                return;
            }

            await EndTaskAsync();
        }

        public string[] GetFilesList(string d, FileInfo[] fiArr, out string ZeroIndEl, string[] FilesList)
        {
            var ext = new List<string> { "nwc" };

            //List forming only from nwc files
            FilesList = Directory.EnumerateFiles(d, "*.*", SearchOption.AllDirectories).
                Where(s => ext.Contains(Path.GetExtension(s).TrimStart((char)46).ToLowerInvariant())).ToArray();

            bool NewFiles = false;
            ZeroIndEl = "";

            //Models by filter exclusion

            foreach (FileInfo f in fiArr)
            {
                if (FileSize != 0 && f.Length < FileSize)
                {
                    FilesList = FilesList.Where(s => s != f.FullName).ToArray();
                }
                if (FileDate != 0 && f.LastWriteTime > DateTime.Now.AddDays(-FileDate))
                {
                    NewFiles = true;
                }
            }

            if (FileDate != 0 && NewFiles == false)
            {
                foreach (FileInfo f in fiArr)
                {
                    FilesList = FilesList.Where(s => s != f.FullName).ToArray();
                }
            }




            //Check if correct models present in directory
            if (FilesList.Length != 0)
            {
                ZeroIndEl = FilesList[0];
                FilesList = FilesList.Where(val => val != FilesList[0]).ToArray();
            }

            return FilesList;
        }

        public void Archiving(string d, string FileNameNWD)
        {
            string ArchiveName = "Архив";
            string ArchiveDir = d + "\\" + ArchiveName;

            // Make a reference to info of a directory.  
            DirectoryInfo di = new DirectoryInfo(d);

            // Get a reference to each file in that directory.
            FileInfo[] fiArr = di.GetFiles();

            //Check directory
            if (fiArr.Length != 0 && Directory.Exists(ArchiveDir) == false)
            {
                Directory.CreateDirectory(ArchiveDir);
            }

            //Archiving
            try
            {
                foreach (FileInfo f in fiArr)
                {
                    if (f.Name.Contains("_Изменения_") == false && f.Name != FileNameNWD && f.Extension == ".nwd")
                    {
                        File.Move(f.FullName, ArchiveDir + "\\" + f.Name.ToString());
                    }
                }
            }

            catch (Exception ex)
            {
                LogWriter(ex.ToString(), CurDateTime);
                return;
            }
        }

        public void LogWriter(string Note, string CurDate)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(LogPath(CurDate), true, System.Text.Encoding.UTF8))
                {
                    sw.WriteLine(Note + "'\n");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        public string LogPath(string CurDateTime)
        {
            string newDirName = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Desktop\NW_Logs\");

            if (Directory.Exists(newDirName) == false)
            {
                Directory.CreateDirectory(newDirName);
            }

            return Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Desktop\NW_Logs\" + CurDateTime + "_Log.txt");
        }

        public void Application_FileInteractiveResolving(object sender, FileInteractiveResolvingEventArgs e)
        {
            //mark as handled, so that Application.FileInteractiveResolving
            if (e is FileInteractiveResolvingEventArgs)
            {
                e.Handled = true;
                //LogWriterResolver(e);
            }
        }
    }


    public partial class FormsInvoke : Form
    {
        public string FolderPathNWCpr { get; private set; }
        public string FolderPathNWDpr { get; private set; }
        public int FileSize { get; private set; }
        public int FileDate { get; private set; }

        public void DialogFormsInvoke()
        {
            //Directories set  
            DS_Form newForm = new DS_Form();

            string FolderPathNWC = newForm.DS_OpenFolderDialogForm("", "Set a directory with input folders with NWC files:").ToString();
            if (FolderPathNWC == "")
            {
                Close();
                return;
            }

            //Property set
            FolderPathNWCpr = FolderPathNWC;

            string FolderPathNWD = newForm.DS_OpenFolderDialogForm("", "Set a directory with output folders for NWD files:").ToString();
            if (FolderPathNWD == "")
            {
                Close();
                return;
            }

            //Property set
            FolderPathNWDpr = FolderPathNWD;
        }

        public FormsInvoke()
        {
            DialogFormsInvoke();
            InitializeComponent();
        }

        private TableLayoutPanel tableLayoutPanel1;
        private TextBox textBox_File_size;
        private CheckBox checkBox_File_size;
        private FlowLayoutPanel flowLayoutPanel2;
        private Button ApplyFilter;
        private Button NoFilter;
        private FlowLayoutPanel flowLayoutPanel_Filters;
        private Label label_File_size;
        private CheckBox checkBox_File_date;
        private TextBox textBox_File_date;
        private Label label_File_date;
        private Button button1;

        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.ApplyFilter = new System.Windows.Forms.Button();
            this.NoFilter = new System.Windows.Forms.Button();
            this.flowLayoutPanel_Filters = new System.Windows.Forms.FlowLayoutPanel();
            this.checkBox_File_size = new System.Windows.Forms.CheckBox();
            this.textBox_File_size = new System.Windows.Forms.TextBox();
            this.label_File_size = new System.Windows.Forms.Label();
            this.checkBox_File_date = new System.Windows.Forms.CheckBox();
            this.textBox_File_date = new System.Windows.Forms.TextBox();
            this.label_File_date = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel2.SuspendLayout();
            this.flowLayoutPanel_Filters.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(0, 0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel_Filters, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(273, 261);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.Controls.Add(this.ApplyFilter);
            this.flowLayoutPanel2.Controls.Add(this.NoFilter);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(3, 211);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(267, 47);
            this.flowLayoutPanel2.TabIndex = 2;
            // 
            // ApplyFilter
            // 
            this.ApplyFilter.Location = new System.Drawing.Point(3, 3);
            this.ApplyFilter.Name = "ApplyFilter";
            this.ApplyFilter.Size = new System.Drawing.Size(111, 42);
            this.ApplyFilter.TabIndex = 0;
            this.ApplyFilter.Text = "ApplyFilter\r\n";
            this.ApplyFilter.UseVisualStyleBackColor = true;
            this.ApplyFilter.Click += new System.EventHandler(this.ApplyFilter_Click);
            // 
            // NoFilter
            // 
            this.NoFilter.Location = new System.Drawing.Point(120, 3);
            this.NoFilter.Name = "NoFilter";
            this.NoFilter.Size = new System.Drawing.Size(141, 42);
            this.NoFilter.TabIndex = 1;
            this.NoFilter.Text = "Continue without filer";
            this.NoFilter.UseVisualStyleBackColor = true;
            this.NoFilter.Click += new System.EventHandler(this.NoFilter_Click_1);
            // 
            // flowLayoutPanel_Filters
            // 
            this.flowLayoutPanel_Filters.Controls.Add(this.checkBox_File_size);
            this.flowLayoutPanel_Filters.Controls.Add(this.textBox_File_size);
            this.flowLayoutPanel_Filters.Controls.Add(this.label_File_size);
            this.flowLayoutPanel_Filters.Controls.Add(this.checkBox_File_date);
            this.flowLayoutPanel_Filters.Controls.Add(this.textBox_File_date);
            this.flowLayoutPanel_Filters.Controls.Add(this.label_File_date);
            this.flowLayoutPanel_Filters.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel_Filters.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flowLayoutPanel_Filters.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel_Filters.Name = "flowLayoutPanel_Filters";
            this.flowLayoutPanel_Filters.Size = new System.Drawing.Size(267, 202);
            this.flowLayoutPanel_Filters.TabIndex = 3;
            // 
            // checkBox_File_size
            // 
            this.checkBox_File_size.AutoSize = true;
            this.checkBox_File_size.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.checkBox_File_size.Location = new System.Drawing.Point(3, 3);
            this.checkBox_File_size.Name = "checkBox_File_size";
            this.checkBox_File_size.Size = new System.Drawing.Size(63, 17);
            this.checkBox_File_size.TabIndex = 2;
            this.checkBox_File_size.Text = "File size";
            this.checkBox_File_size.UseVisualStyleBackColor = true;
            this.checkBox_File_size.CheckedChanged += new System.EventHandler(this.checkBox_File_size_CheckedChanged);
            // 
            // textBox_File_size
            // 
            this.textBox_File_size.Location = new System.Drawing.Point(3, 26);
            this.textBox_File_size.Name = "textBox_File_size";
            this.textBox_File_size.Size = new System.Drawing.Size(100, 20);
            this.textBox_File_size.TabIndex = 3;
            this.textBox_File_size.Tag = "";
            this.textBox_File_size.Visible = false;
            this.textBox_File_size.TextChanged += new System.EventHandler(this.textBox_File_size_TextChanged);
            // 
            // label_File_size
            // 
            this.label_File_size.AutoSize = true;
            this.label_File_size.Location = new System.Drawing.Point(3, 49);
            this.label_File_size.Name = "label_File_size";
            this.label_File_size.Size = new System.Drawing.Size(168, 13);
            this.label_File_size.TabIndex = 4;
            this.label_File_size.Text = "Enter the smallest size of a file, KB";
            this.label_File_size.Visible = false;
            // 
            // checkBox_File_date
            // 
            this.checkBox_File_date.AutoSize = true;
            this.checkBox_File_date.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.checkBox_File_date.Location = new System.Drawing.Point(3, 65);
            this.checkBox_File_date.Name = "checkBox_File_date";
            this.checkBox_File_date.Size = new System.Drawing.Size(66, 17);
            this.checkBox_File_date.TabIndex = 5;
            this.checkBox_File_date.Text = "File date";
            this.checkBox_File_date.UseVisualStyleBackColor = true;
            this.checkBox_File_date.CheckedChanged += new System.EventHandler(this.checkBox_File_date_CheckedChanged);
            // 
            // textBox_File_date
            // 
            this.textBox_File_date.Location = new System.Drawing.Point(3, 88);
            this.textBox_File_date.Name = "textBox_File_date";
            this.textBox_File_date.Size = new System.Drawing.Size(100, 20);
            this.textBox_File_date.TabIndex = 6;
            this.textBox_File_date.Tag = "";
            this.textBox_File_date.Visible = false;
            this.textBox_File_date.TextChanged += new System.EventHandler(this.textBox_File_date_TextChanged);
            // 
            // label_File_date
            // 
            this.label_File_date.AutoSize = true;
            this.label_File_date.Location = new System.Drawing.Point(3, 111);
            this.label_File_date.Name = "label_File_date";
            this.label_File_date.Size = new System.Drawing.Size(235, 26);
            this.label_File_date.TabIndex = 7;
            this.label_File_date.Text = "Enter the last file write date (in days from current date).";
            this.label_File_date.Visible = false;
            // 
            // Filters_Applying
            // 
            this.ClientSize = new System.Drawing.Size(273, 261);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.button1);
            this.Name = "Filters_Applying";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Filters applying";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel2.ResumeLayout(false);
            this.flowLayoutPanel_Filters.ResumeLayout(false);
            this.flowLayoutPanel_Filters.PerformLayout();
            this.ResumeLayout(false);

        }

        public void ApplyFilter_Click(object sender, EventArgs e)
        {

            if (checkBox_File_size.Checked && textBox_File_size.Text == "")
            {
                MessageBox.Show("Enter any value to the field");
                return;
            }

            if (checkBox_File_date.Checked && textBox_File_date.Text == "")
            {
                MessageBox.Show("Enter any value to the field");
                return;
            }

            this.Close();

            NWC_Assembly_Plugin nWC_Assembly_Plugin = new NWC_Assembly_Plugin();
            nWC_Assembly_Plugin.MainProgram(FolderPathNWCpr, FolderPathNWDpr, FileSize, FileDate);
        }

        private void NoFilter_Click_1(object sender, EventArgs e)
        {
            this.Close();
            NWC_Assembly_Plugin nWC_Assembly_Plugin = new NWC_Assembly_Plugin();
            nWC_Assembly_Plugin.MainProgram(FolderPathNWCpr, FolderPathNWDpr, FileSize, FileDate);
        }

        private void checkBox_File_size_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_File_size.Checked)
            {
                textBox_File_size.Visible = true;
                label_File_size.Visible = true;
            }
            else
            {
                FileSize = 0;
                textBox_File_size.Clear();
                textBox_File_size.Visible = false;
                label_File_size.Visible = false;
            }

        }

        private void checkBox_File_date_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_File_date.Checked)
            {
                textBox_File_date.Visible = true;
                label_File_date.Visible = true;
            }
            else
            {
                FileDate = 0;
                textBox_File_date.Text = "";
                textBox_File_date.Visible = false;
                label_File_date.Visible = false;
            }

        }

        private void textBox_File_size_TextChanged(object sender, EventArgs e)
        {
            if (!int.TryParse(textBox_File_size.Text, out int ParsedValue) && textBox_File_size.Text != "")
            {
                MessageBox.Show("This is a number only field");
                return;
            }
            FileSize = ParsedValue * 1000;
        }

        private void textBox_File_date_TextChanged(object sender, EventArgs e)
        {
            {
                if (!int.TryParse(textBox_File_date.Text, out int ParsedValue) && textBox_File_date.Text != "")
                {
                    MessageBox.Show("This is a number only field");
                    return;
                }

                FileDate = ParsedValue;
            }
        }


    }
}
