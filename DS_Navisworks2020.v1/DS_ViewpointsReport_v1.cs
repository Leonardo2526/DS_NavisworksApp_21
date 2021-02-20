using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;

//COM spaces
using Autodesk.Navisworks.Api.Interop.ComApi;
using Autodesk.Navisworks.Api.Plugins;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

//User spaces
using DS_Space;


namespace DS_NWClass_ViewpointsReport
{
    /*
    DS_Output dS_Output = new DS_Output();
    dS_Output.DS_StreamWriter(test.DisplayName);
    */

    [Plugin("DS_ClashVPReport_v1", "DS", ToolTip = "Viewpoints report creation from clash tests results.", DisplayName = "DS_ClashVPReport_v1")]

    public class ViewpointsReport : AddInPlugin
    {

        //Get current date and time 
        public string CurDate = DateTime.Now.ToString("yyMMdd");

        private static string SourseFolderPath;

        List<string> SavedFiles = new List<string>();
        List<string> FilesWithNoClashTests = new List<string>();
        List<string> FilesWithNoClashRes = new List<string>();

        public override int Execute(params string[] parameters)
        {
            Program();

            return 0;
        }

        public void Program()
        {
            //New form creation
            DS_Form newForm = new DS_Form();

            string FolderPathNWD = newForm.DS_OpenFolderDialogForm("", "Set directory for projects with NWF files:").ToString();
            if (FolderPathNWD == "")
            {
                newForm.Close();
                return;
            }
            SourseFolderPath = FolderPathNWD;
            string CurDateTime = DateTime.Now.ToString("yyMMdd_HHmmss");

            DS_Output dS_Output = new DS_Output
            {
                DS_WritePath = SourseFolderPath + "\\" + "Log_ClashVPReport_" + CurDateTime + ".txt"
            };

            SavedFiles.Clear();
            SavedFiles.Add("Files has been saved: " + "\n");
            FilesWithNoClashTests.Clear();
            FilesWithNoClashTests.Add("No clash tests exist in files: " + "\n");
            FilesWithNoClashRes.Clear();
            FilesWithNoClashRes.Add("No clashes are in files: " + "\n");

            FilesCheck(SourseFolderPath);

            DirIterate(SourseFolderPath);


            //Start log writing
            if (SavedFiles.Count > 1)
            {
                SavedFiles.Add("\n");
                foreach (string s in SavedFiles)
                {
                    dS_Output.DS_StreamWriter(s);
                }
            }
            else
            {
                dS_Output.DS_StreamWriter("No NWF files have been found in this folder.");
                MessageBox.Show("Errors occurred!" + "\n" + "Look at " + dS_Output.DS_WritePath + " for details.");
                return;
            }

            if (FilesWithNoClashTests.Count > 1)
            {
                FilesWithNoClashTests.Add("\n");
                foreach (string s in FilesWithNoClashTests)
                    dS_Output.DS_StreamWriter(s);
            }

            if (FilesWithNoClashRes.Count > 1)
            {
                FilesWithNoClashRes.Add("\n");
                foreach (string s in FilesWithNoClashRes)
                    dS_Output.DS_StreamWriter(s);
            }

            MessageBox.Show("Process complete! See report: " + "\n" + dS_Output.DS_WritePath);
        }

        public void DirIterate(string CheckPath)
        {
            try
            {
                //Check folders
                foreach (string d in Directory.GetDirectories(CheckPath))
                {
                    FilesCheck(d);
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
        }

        public void FilesCheck(string PathCheck)
        {
            //Tools for work with files creation
            DS_DirTools dS_DirTools = new DS_DirTools();

            bool fe = dS_DirTools.DirCheckForFiles(PathCheck, out string[] FilesList, "nwf");
            if (fe == false)
            {
                return;
            }

            //Through each file iterating 
            foreach (string f in FilesList)
            {
                //Open Navisworks document
                Navisworks_Tools navisworks_Tools = new Navisworks_Tools(f);
                Document oDoc = navisworks_Tools.OpenDoc();
                DocumentClash documentClash = oDoc.GetClash();
                DocumentClashTests oDCT = documentClash.TestsData;

                if (oDCT.Tests.Count != 0)
                {
                    oDCT.TestsRunAllTests();

                    //Through each clash test iterating
                    Clash_Viewpoints clash_Viewpoints = new Clash_Viewpoints(oDoc);

                    clash_Viewpoints.CheckHomeView();

                    clash_Viewpoints.ClashElementSearch();
                    clash_Viewpoints.ViewpointsCreation(out bool ClashTestsResultsExist);

                    if (ClashTestsResultsExist == false)
                    {
                        FilesWithNoClashRes.Add(f);
                    }
                    else
                    {
                        //Set current view
                        DS_NW_Viewpoint_tools dS_NW_Viewpoint_Tools = new DS_NW_Viewpoint_tools();
                        dS_NW_Viewpoint_Tools.HomeViewpoint_Set(out _);

                        navisworks_Tools.FilesSave();

                        SavedFiles.Add(f);
                    }
                    
                }
                else
                FilesWithNoClashTests.Add(f);
            }

        }


    }

    public class Navisworks_Tools : ViewpointsReport
    {
        private string FilePath;
        private Document oDoc;

        public Navisworks_Tools(string filename)
        {
            FilePath = filename;
        }

        public Document OpenDoc()
        {
            Autodesk.Navisworks.Api.Application.FileInteractiveResolving += Application_FileInteractiveResolving;

            Autodesk.Navisworks.Api.Application.ActiveDocument.TryOpenFile(FilePath);
            oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            return oDoc;
        }

        private void Application_FileInteractiveResolving(object sender, FileInteractiveResolvingEventArgs e)
        {
            //mark as handled, so that Application.FileInteractiveResolving
            if (e is FileInteractiveResolvingEventArgs)
            {
                e.Handled = true;
            }
        }

        public void FilesSave()
        {
            //Save NWF file
            oDoc.SaveFile(FilePath);

            //NWD file name
            string NWDName = Path.GetFileNameWithoutExtension(FilePath) + "_" + CurDate + ".nwd";

            //NWD file path
            string NWDPath = Path.GetDirectoryName(FilePath) + "\\" + CurDate;

            if (Directory.Exists(NWDPath) == false)
            {
                Directory.CreateDirectory(NWDPath);
            }

            //Save NWD file
            oDoc.SaveFile(NWDPath + "\\" + NWDName);
        }


    }

    public class Clash_Viewpoints
    {
        readonly InwOpState oState = ComApiBridge.State;
        InwOpClashElement m_clash = null;
        static private Document oDoc;

        public Clash_Viewpoints(Document odoc)
        {
            oDoc = odoc;
        }


        public void ClashElementSearch()
        {
            //find the clash detective plugin
            foreach (InwBase oPlugin in oState.Plugins())
            {
                if (oPlugin.ObjectName == "nwOpClashElement")
                {
                    m_clash = (InwOpClashElement)oPlugin;
                    break;
                }
            }

            if (m_clash == null)
            {
                MessageBox.Show("cannot find clash test plugin!");
                return;
            }
        }

        public void ViewpointsCreation(out bool ClashTestsResultsExist)
        {
            InwOpGroupView TopGroup = null;
            ProjectionChange_Ort();
            ChangeModelColor();

            ClashTestsResultsExist = false;

            try
            {
                foreach (InwOclClashTest clashTest in m_clash.Tests())
                {
                    clashTest.status = nwEClashTestStatus.eClashTestStatus_OK;

                    if (clashTest.results().Count == 0)
                    {
                        continue;
                    }

                    ClashTestsResultsExist = true;

                    if (TopGroup == null)
                        TopGroup = TopGroupCreate();

                    InwOpGroupView SubGroup = SubGroupCreate(clashTest, TopGroup);

                    // get the first Test and its first clash result 
                    foreach (InwOclTestResult clashResult in clashTest.results())
                    {

                        if (clashResult.status == nwETestResultStatus.eTestResultStatus_ACTIVE |
                            clashResult.status == nwETestResultStatus.eTestResultStatus_NEW)

                        {
                            ModelItem oItem1 = null;
                            ModelItem oItem2 = null;

                            //Get clashed items                      
                            oItem1 = ComApiBridge.ToModelItem(clashResult.Path1);
                            oItem2 = ComApiBridge.ToModelItem(clashResult.Path2);

                            //Put clashed items to model collection 
                            ModelItemCollection ClashedItems = new ModelItemCollection
                            {
                                oItem1,
                                oItem2
                            };

                            //Overwrite items color
                            oDoc.Models.OverridePermanentColor(new ModelItem[1] { ClashedItems.ElementAtOrDefault(0) },
                                Autodesk.Navisworks.Api.Color.Red);
                            oDoc.Models.OverridePermanentColor(new ModelItem[1] { ClashedItems.ElementAtOrDefault(1) },
                                Autodesk.Navisworks.Api.Color.Green);

                            // Create a copy of current selection
                            ComApi.InwOpSelection ComClashSelection = ComApiBridge.ToInwOpSelection(ClashedItems);

                            // Create a new empty selection
                            InwOpSelection2 AllModel = oState.ObjectFactory(nwEObjectType.eObjectType_nwOpSelection, null, null) as InwOpSelection2;

                            // Get the new selection to contain the entire model
                            AllModel.SelectAll();

                            // Subtract the current selection, so it contains the unselected part of model
                            AllModel.SubtractContents(ComClashSelection);

                            // create a temporary saved viewpoint 
                            InwOpView oSV = oState.ObjectFactory(nwEObjectType.eObjectType_nwOpView);

                            //Apply attributes
                            oSV.ApplyHideAttribs = true;
                            oSV.ApplyMaterialAttribs = true;

                            oSV.name = clashResult.name;

                            //Get and save viewpoint
                            oSV.anonview.ViewPoint = clashResult.GetSuitableViewPoint().Copy();

                            SubGroup.SavedViews().Add(oSV);

                            oDoc.Models.OverridePermanentColor(new ModelItem[1] { ClashedItems.ElementAtOrDefault(0) },
                                Autodesk.Navisworks.Api.Color.FromByteRGB(63,63,63));
                            oDoc.Models.OverridePermanentColor(new ModelItem[1] { ClashedItems.ElementAtOrDefault(1) },
                                Autodesk.Navisworks.Api.Color.FromByteRGB(63, 63, 63));

                            oDoc.CurrentSelection.Clear();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (TopGroup == null)
                return;
            
            oState.SavedViews().Add(TopGroup);
        }

        public InwOpGroupView TopGroupCreate()
        {
            ViewpointsReport viewpointsReport = new ViewpointsReport();

            // create a temporary saved viewpoint
            InwOpFolderView oSF = oState.ObjectFactory(nwEObjectType.eObjectType_nwOpFolderView);
            oSF.name = viewpointsReport.CurDate;

            //Folder assigning as group
            InwOpGroupView TopGroup = oSF;

            return TopGroup;
        }

        public InwOpGroupView SubGroupCreate(InwOclClashTest clashTest, InwOpGroupView TopGroup)
        {
            // create a temporary saved subfolder
            InwOpFolderView oSsF = oState.ObjectFactory(nwEObjectType.eObjectType_nwOpFolderView);
            oSsF.name = clashTest.name;

            //Add subfolder as item of the Top group
            TopGroup.SavedViews().Add(oSsF);

            //Subfolder assigning as group
            InwOpGroupView SubGroup = oSsF;

            return SubGroup;
        }

        public void ProjectionChange_Ort()
        {
            Viewpoint oVP = Autodesk.Navisworks.Api.Application.ActiveDocument.CurrentViewpoint.CreateCopy();
            oVP.Projection = ViewpointProjection.Orthographic;
            oVP.RenderStyle = ViewpointRenderStyle.Wireframe;
            Autodesk.Navisworks.Api.Application.ActiveDocument.CurrentViewpoint.CopyFrom(oVP);
        }

        public void CheckHomeView()
        {
            bool HVExist = false;

            try
            {
                foreach (SavedItem oSVP in oDoc.SavedViewpoints.Value)
                {
                    if (oSVP.DisplayName == "HomeView")
                    {
                        HVExist = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (HVExist == false)
            {
                DS_NW_Viewpoint_tools dS_NW_Viewpoint_Tools = new DS_NW_Viewpoint_tools();
                dS_NW_Viewpoint_Tools.HomeViewpoint_Set(out InwOpView oSV);

                //Save viewpoint
                oSV.name = "HomeView";
                oState.SavedViews().Add(oSV);
            }
        }

        public void ChangeModelColor()
        {
            int i = 0;
            Document oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            foreach (ModelItem modelItem in oDoc.Models.CreateCollectionFromRootItems())
            {
                IEnumerable<ModelItem> items = oDoc.Models[i].RootItem.Descendants;
                oDoc.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.FromByteRGB(63, 63, 63));
                i += 1;
            }
        }
    }
}
