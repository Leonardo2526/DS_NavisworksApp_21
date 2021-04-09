using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Security;
using System.Security.AccessControl;
using System.IO;
using System.Diagnostics;

using Autodesk.Navisworks.Api;

namespace DS_Space
{
   
    public class DS_Form : Form
        {
            private Button selectButton;
            private OpenFileDialog openFileDialog1;

            public string DS_OpenFileDialogForm_txt()
            {
                openFileDialog1 = new OpenFileDialog()
                {
                    FileName = "Select a text file",
                    Filter = "Text files (*.txt)|*.txt",
                    Title = "Open text file"
                };

                selectButton = new Button()
                {
                    Size = new Size(100, 20),
                    Location = new Point(15, 15),
                    Text = "Select file"
                };

                string filePath = "";

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    SelectButton_Click(ref filePath);
                }
                return filePath;
            }

        public string DS_OpenFileDialogForm(string filename = "Select a file")
        {
                openFileDialog1 = new OpenFileDialog()
                {
                    FileName = filename,
                    Title = "Open file",
                    
                };

        selectButton = new Button()
            {
                Size = new Size(100, 20),
                Location = new Point(15, 15),
                Text = "Select file"
            };

            string filePath = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SelectButton_Click(ref filePath);
            }
            return filePath;
        }
        private string SelectButton_Click(ref string filePath)
            {
                try
                {
                    filePath = openFileDialog1.FileName.ToString();
                }
                catch (SecurityException ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
                return filePath;
            }


        public string DS_OpenFolderDialogForm(string folderPath = "", string description = "Chose folder")
        {
            

                FolderBrowserDialog fbd;
         
                if (folderPath != "")
                    {
                        fbd = new FolderBrowserDialog
                        {
                            Description = description,
                            SelectedPath = folderPath
                        };
                    }
                    else
                    {
                        fbd = new FolderBrowserDialog
                        {
                            Description = description
                        };
                    }

                    // Show testDialog as a modal dialog
                    DialogResult result = fbd.ShowDialog();
                    string sfp = fbd.SelectedPath;

            
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        if (DS_HasWritePermissionOnDir(sfp) == true)
                        {
                            return sfp;
                        }
                        else
                        {
                            MessageBox.Show("Error access to path!");
                            return "";
                        }
                    }
            

             return "";
        }
        public bool DS_HasWritePermissionOnDir(string path)
        {
            //Check directory permissions
            var writeAllow = false;
            var writeDeny = false;
            var accessControlList = Directory.GetAccessControl(path);
            if (accessControlList == null)
                return false;
            var accessRules = accessControlList.GetAccessRules(true, true, typeof(System.Security.Principal.SecurityIdentifier));
            if (accessRules == null)
                return false;

            //Check rules
            foreach (FileSystemAccessRule rule in accessRules)
            {
                if ((FileSystemRights.Write & rule.FileSystemRights) != FileSystemRights.Write) continue;

                if (rule.AccessControlType == AccessControlType.Allow)
                    writeAllow = true;
                else if (rule.AccessControlType == AccessControlType.Deny)
                    writeDeny = true;
            }
            return writeAllow && !writeDeny;
        }

    }

    public class DS_Output
    {
        public string DS_WritePath { get; set; } = @"C:\Users\dnknt\OneDrive\Рабочий стол\Output.txt";

        public void DS_StreamWriter(string OutputTxt, bool Add = true)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(DS_WritePath, Add, Encoding.UTF8))
                {
                    sw.WriteLine(OutputTxt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            };
            return;

        }

        public void DS_ResultMessage()
        {
            if (Directory.Exists(DS_WritePath))
            {
                MessageBox.Show("Output file has been saved to: " + DS_WritePath);
            }
        }

        public string DS_PathNameCreate(string CurDateTime, string DirName = @"%USERPROFILE%\Desktop\NW_Logs\", string FileName = "Log", string FileExt = ".txt")
        {
            string ExpDirName = Environment.ExpandEnvironmentVariables(DirName);

            if (Directory.Exists(ExpDirName) == false)
            {
                Directory.CreateDirectory(ExpDirName);
            }

            return Environment.ExpandEnvironmentVariables(DirName + CurDateTime + "_" + FileName + FileExt);
        }
    }

    public class DS_DirTools
    {
        public bool DirCheckForFiles(string dir, out string[] FilesList, string ext1 = "", string ext2 = "")
            //Check top directory for files existance. Without subdirectories.
        {
            // Get a reference to each file in that directory.
            FilesList = Directory.GetFiles(dir);

            //Check files for extension
            if (ext1 != "" | ext2 != "")
                {
                //Extensions list
                var FileExt = new List<string> { ext1, ext2 };

                //Array forming from files with "ext" extension only
                FilesList = Directory.EnumerateFiles(dir, "*.*", SearchOption.TopDirectoryOnly).
                    Where(s => FileExt.Contains(Path.GetExtension(s).TrimStart((char)46).ToLowerInvariant())).ToArray();
                }

            if (FilesList.Length == 0)
                return false;
            else
                return true;
        }
    }

    public class DS_NW_Viewpoint_tools
    {
        public void CurentViewpoint_Save()
        //Get current viewpoint and save it to items
        {
            Document oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            // Create the saved viewpoint with the new viewpoint
            SavedViewpoint oNewVP = new SavedViewpoint(oDoc.CurrentViewpoint.ToViewpoint());

            oNewVP.DisplayName = "MySavedView";

            //Add viewpoint to saved items
            oDoc.SavedViewpoints.AddCopy(oNewVP);
        }

        public void SetCurrentViewpointFromSaved()
        //Set current view from specific saved viewpoint
        {
            Document oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            //Through viewpoints iterating
            foreach (SavedViewpoint oSVP in oDoc.SavedViewpoints.Value)
            {
                if (oSVP.DisplayName == "MySavedView")
                {
                    //Set current viewpoint to oSVP
                    oDoc.CurrentViewpoint.CopyFrom(oSVP.Viewpoint);
                }
            }
        }

        public void HomeViewpoint_IfSet()
        //Get home viewpoint (if it was set) and save it to items
        {
            Document oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            try
            {
                // Create the saved viewpoint with the new viewpoint
                SavedViewpoint oNewVP = new SavedViewpoint(oDoc.HomeView);

                oNewVP.DisplayName = "MyHomeView";

                //Add viewpoint to saved items
                oDoc.SavedViewpoints.AddCopy(oNewVP);

                //Set current viewpoint to oNewVP
                oDoc.CurrentViewpoint.CopyFrom(oNewVP.Viewpoint);
            }
            catch (Exception e)
            {
                MessageBox.Show("No set home view in model!" + "\n\n" + e);
            };

        }
      
       
    }


}

