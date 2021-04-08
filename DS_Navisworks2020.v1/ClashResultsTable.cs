using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Security.AccessControl;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;



namespace DS_ClashReport
{
    [Plugin("DS_ClashResultsTable_v4.1", "DS", ToolTip = "Clash results status output", DisplayName = "DS_ClashResultsTable_v4.1")]

    public class ClashResultsTable : AddInPlugin
    {

        public override int Execute(params string[] parameters)
        {

            Program();

            return 0;
        }


        public void Program()
        {
            //Name of input file
            OpenFileDialogForm_1 opFile = new OpenFileDialogForm_1();
            string FilePathTxt = opFile.OpenFileDialogForm().ToString();
            if (FilePathTxt == "")
            {
                opFile.Close();
                return;
            }

            string newDirName = "";

            //Files pathes reading
            string[] lines = File.ReadAllLines(FilePathTxt).Where(s => s.Trim() != string.Empty).ToArray();
            string CurDate = DateTime.Now.ToString("yyMMdd");
            string CurDateTime = DateTime.Now.ToString("yyMMdd_HHmmss");

            if (Testing(lines, CurDateTime) == 0)
            {
                //Output folder for Excel
                if (newDirName == "")
                {
                    OpenFolder opFold = new OpenFolder();
                    newDirName = opFold.OpenFolderDialogForm() + ((char)92).ToString();
                    if (newDirName == "")
                    {
                        return;
                    }
                }


                string FileName = CurDate + "_Сводный_отчёт_по_коллизиям.xlsx";

                ExcelCheckActiveWorkbook(FileName);

                //Open Excel
                var excelApp = new Excel.Application
                {
                    Visible = false
                };
                excelApp.Workbooks.Add();
                Excel.Worksheet newWorksheet = null;
                Excel.Sheets sheets = excelApp.ActiveWorkbook.Sheets;
                Excel.Worksheet lastSheet = sheets[1];

                //Intiating main process
                try
                {
                    for (int il = 0; il < lines.Length; il++)
                    {
                        string FilePath = lines[il];
                        string GetFileName = Path.GetFileName(FilePath);

                        string WorksheetName = NameOfSheet(GetFileName, out string field1, out string field2);

                        Autodesk.Navisworks.Api.Application.FileInteractiveResolving += Application_FileInteractiveResolving;

                        Autodesk.Navisworks.Api.Application.ActiveDocument.TryOpenFile(FilePath);

                        Document oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                        DocumentClash documentClash = oDoc.GetClash();
                        DocumentClashTests oDCT = documentClash.TestsData;


                        if (oDCT.Tests.Count != 0)
                        {
                            TestSearchResultsInitialising(oDCT, WorksheetName, ref newWorksheet, ref lastSheet, sheets);
                        }
                        else
                        {
                            LogWriter("No clash tests found in the file: ", FilePath, CurDateTime);
                            MessageBox.Show("Process has been stoped because errors occured!" + "\n" + "Log saved: " + LogPath(CurDateTime));
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }


                if (sheets.Count > 1)
                    sheets[1].Delete();

                //Total sheet creating
                TotalSheet(excelApp, newWorksheet, lines);

                try
                {
                    excelApp.ActiveWorkbook.SaveAs(newDirName + FileName);
                }
                catch
                { }

                excelApp.Visible = true;
                MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow, "Done!");

            }
            else
            {
                MessageBox.Show("Process has been stoped because errors occured!" + "\n" + "Log saved: " + LogPath(CurDateTime));
                return;
            }

        }

        public void TestSearchResultsInitialising(DocumentClashTests oDCT, string WorksheetName, ref Excel.Worksheet newWorksheet, ref Excel.Worksheet lastSheet, Excel.Sheets sheets)
        {
            //Handling with external referencies

            newWorksheet = sheets.Add(Type.Missing, lastSheet, Type.Missing, Type.Missing);
            newWorksheet.Name = WorksheetName;
            lastSheet = newWorksheet;

            SheetsHeadings(newWorksheet);

            ClashTestsSearch(oDCT, newWorksheet);

            //Columns AutoFit
            for (int i = 1; i <= 8; i++)
            {
                newWorksheet.Columns[i].AutoFit();
            }

            int sheetInd = newWorksheet.Index;
            LastCellSearch(sheets, sheetInd, out int lastRow, out int lastColumn);

            ExcelSheetsFormat(newWorksheet, lastRow, lastColumn);


        }

        public void SheetsHeadings(Excel.Worksheet newWorksheet)
        {
            // Establish column headings in cells.
            List<string> listOfHeadings = new List<string>()
            {
                "№",
                "Имя проверки",
                "Всего",
                "Новые",
                "Активные",
                "Проверенные",
                "Подтвержденные",
                "Исправленные"
            };

            int i = 0;
            foreach (string head in listOfHeadings)
            {
                i += 1;
                newWorksheet.Cells[1, i] = head;
            }
        }
        public void TotalSheetsHeadings(Excel.Worksheet newWorksheet)
        {
            // Establish column headings in cells.
            List<string> listOfHeadings = new List<string>()
            {
                "№ на сит. плане",
                "Объект",
                "Всего",
                "Новые",
                "Активные",
                "Проверенные",
                "Подтвержденные",
                "Исправленные"
            };

            int i = 0;
            foreach (string head in listOfHeadings)
            {
                i += 1;
                newWorksheet.Cells[1, i] = head;
            }
        }

        public void ExcelCheckActiveWorkbook(string excelFilePath)
        {
            try
            {
                Excel.Application excelApp =
                    (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                if (excelApp.ActiveWorkbook.FullName == excelFilePath)
                {
                    excelApp.Workbooks.Close();
                    excelApp.Quit();
                }
            }
            catch
            {
            }
        }

        public void ClashTestsSearch(DocumentClashTests oDCT, Excel.Worksheet newWorksheet)
        {
            //ClashTests names output
            int i;
            int ic = 0;
            int[] Sum = new int[6];

            //Clash tests iterating
            foreach (ClashTest ctest in oDCT.Tests)
            {
                newWorksheet.Cells[ic + 2, 1] = ic + 1;
                newWorksheet.Cells[ic + 2, 2] = ctest.DisplayName;
                int[] StatusNum = new int[6];

                //ClashTestsResults properties output
                foreach (SavedItem issue in ctest.Children)
                {
                    ClashResult clash = issue as ClashResult;
                    if (clash != null)
                    {
                        if (clash.Status == ClashResultStatus.New)
                            StatusNum[1] = StatusNum[1] + 1;
                        else if (clash.Status == ClashResultStatus.Active)
                            StatusNum[2] = StatusNum[2] + 1;
                        else if (clash.Status == ClashResultStatus.Reviewed)
                            StatusNum[3] = StatusNum[3] + 1;
                        else if (clash.Status == ClashResultStatus.Approved)
                            StatusNum[4] = StatusNum[4] + 1;
                        else if (clash.Status == ClashResultStatus.Resolved)
                            StatusNum[5] = StatusNum[5] + 1;
                    }
                }

                //Total clash test results ammountB
                StatusNum[0] = StatusNum[1] + StatusNum[2] + StatusNum[3] + StatusNum[4] + StatusNum[5];

                //Data assigning to Excel cells
                for (i = 0; i <= 5; i++)
                {
                    newWorksheet.Cells[ic + 2, 3 + i] = StatusNum[i];
                    Sum[i] += StatusNum[i];
                }
                ic += 1;
            }

            newWorksheet.Cells[ic + 2, 2] = "ИТОГО:";
            for (i = 0; i <= 5; i++)
            {
                newWorksheet.Cells[ic + 2, 3 + i] = Sum[i];
            }
        }

        private void TotalSheet(Excel.Application excelApp, Excel.Worksheet newWorksheet, string[] lines)
        {
            Excel.Sheets sheets = excelApp.ActiveWorkbook.Sheets;
            newWorksheet = sheets.Add(Type.Missing, newWorksheet, Type.Missing, Type.Missing);
            newWorksheet.Name = "Сводный";

            TotalSheetsHeadings(newWorksheet);

            //Filling two first columns
            for (int il = 0; il < lines.Length; il++)
            {
                string FilePath = lines[il];
                string GetFileName = Path.GetFileName(FilePath);

                Excel.Range range = newWorksheet.get_Range("A1", "A2");
                range.EntireColumn.NumberFormat = "@";

                NameOfSheet(GetFileName, out string field1, out string field2);
                field1 = field1.Replace("_", "");
                newWorksheet.Cells[il + 2, 1] = field1;
                newWorksheet.Cells[il + 2, 2] = field2;
            }

            //Sheets iterating
            LastCellSearch(sheets, 1, out int lastRow, out int lastColumn);
            double[] Sum = new double[lastColumn + 1];
            int j;
            int sheetCnt = 0;
            foreach (Excel.Worksheet sheet in sheets)
            {
                if (sheet.Name != "Сводный")
                {
                    LastCellSearch(sheets, sheet.Index, out lastRow, out lastColumn);

                    sheetCnt += 1;
                    //MessageBox.Show(sheetCnt.ToString() + "_"+ sheet);
                    try
                    {
                        for (j = 3; j <= lastColumn; j++)
                        {
                            newWorksheet.Cells[sheetCnt + 1, j] = sheet.Cells[lastRow, j];
                            Sum[j] = Sum[j] + sheet.Cells[lastRow, j].Value;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    lastRow = sheetCnt + 1;
                }
            }

            int sheetInd = newWorksheet.Index;

            //Data record to Excel cells
            newWorksheet.Cells[lastRow + 1, 2] = "ИТОГО:";
            for (j = 3; j <= lastColumn; j++)
            {
                newWorksheet.Cells[lastRow + 1, j] = Sum[j];
            }

            //Columns autofit;
            for (j = 1; j <= lastColumn; j++)
            {
                newWorksheet.Columns[j].AutoFit();
            }

            LastCellSearch(sheets, sheetInd, out lastRow, out lastColumn);
            ExcelSheetsFormat(newWorksheet, lastRow, lastColumn);
        }

        public void LastCellSearch(Excel.Sheets sheets, int SheetInd, out int lastRow, out int lastColumn)
        {
            //Get last colomn and row number
            Excel.Range lastCell = sheets[SheetInd].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            lastColumn = lastCell.Column;
            lastRow = lastCell.Row;
        }

        public void ExcelSheetsFormat(Excel.Worksheet newWorksheet, int rowNumb, int colNumb)
        {
            //Excel sheet formating
            newWorksheet.Cells[rowNumb, colNumb].EntireRow.Font.Bold = true;

            Excel.Range range = newWorksheet.Range[newWorksheet.Cells[1, 1], newWorksheet.Cells[rowNumb, colNumb]];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.Weight = Excel.XlBorderWeight.xlThin;
        }

        public string NameOfSheet(string GetFileName, out string field1, out string field2)
        {
            field1 = "";
            field2 = "";

            if (GetFileName.Contains(".nwd") == true || GetFileName.Contains(".nwf") == true)
            {
                int i = 0;
                int indS = 0;
                int[] n = new int[GetFileName.Length];
                int[] ind = new int[GetFileName.Length];

                //Through each symbol in file name iterating
                foreach (char sign in GetFileName)
                {
                    indS += 1;

                    if (sign.ToString() == "_")
                    {
                        i += 1;

                        //Count of findings
                        n[i] = i;

                        //Index of findings
                        ind[i] = indS;

                        //Record of new string
                        if (n[i] == 2)
                        {
                            field1 += GetFileName.Substring(ind[i - 1], ind[i] - ind[i - 1]);
                        }
                        else if (n[i] == 5)
                        {
                            field2 += GetFileName.Substring(ind[i - 1], ind[i] - ind[i - 1] - 1);
                        }
                    }
                }
            }

            return field1 + field2;

        }

        private void Application_FileInteractiveResolving(object sender, FileInteractiveResolvingEventArgs e)
        {
            //mark as handled, so that Application.FileInteractiveResolving
            if (e is FileInteractiveResolvingEventArgs)
            {
                e.Handled = true;
                //LogWriterResolver(e);
            }
        }


        public void LogWriterResolver(FileInteractiveResolvingEventArgs e, string CurDateTime)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(LogPath(CurDateTime), true, System.Text.Encoding.UTF8))
                {
                    sw.WriteLine("Can't resolve external reference: '" + e.FileReference + "'. File path: " + e.ReferringFileName + "'\n");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }


        public void LogWriter(string Note, string FilePath, string CurDate)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(LogPath(CurDate), true, System.Text.Encoding.UTF8))
                {
                    sw.WriteLine(Note + FilePath + "'\n");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        public int Testing(string[] lines, string CurDate)
        {
            int error = 0;
            string[] WorksheetName = new string[lines.Length];
            try
            {
                for (int il = 0; il < lines.Length; il++)
                {
                    if (lines[il].Contains(((char)34).ToString()))
                    {
                        lines[il] = lines[il].Replace(((char)34).ToString(), "");
                    }

                    string FilePath = lines[il];
                    string GetFileName = Path.GetFileName(FilePath);
                    WorksheetName[il] = NameOfSheet(GetFileName, out string field1, out string field2);

                    if (File.Exists(lines[il]) == false)
                    {
                        LogWriter("No such file path: '", FilePath, CurDate);
                        error = 2;

                        //Output if array contain repeated elements
                        if (IFDuplicateKeys(WorksheetName) == true)
                        {
                            LogWriter("Duplicated sheets names '" + WorksheetName[il] + "'. Rename file's fields #2 and #4: '", FilePath, CurDate);
                            error = 1;
                        }
                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow, ex.ToString());
            }
            return error;
        }

        private string LogPath(string CurDateTime)
        {
            string newDirName = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Desktop\NW_Logs\");

            if (Directory.Exists(newDirName) == false)
            {
                Directory.CreateDirectory(newDirName);
            }

            return Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Desktop\NW_Logs\" + CurDateTime + "_Log.txt");
        }

        public bool IFDuplicateKeys(string[] array)
        {
            //Detection if array contain repeates elements
            var duplKeys = array.GroupBy(x => x)
                            .Where(group => group.Count() > 1)
                            .Select(group => group.Key);
            if (duplKeys.Count() != 0)
            {
                return true;
            }
            else
                return false;
        }



    }

    public class OpenFileDialogForm_1 : Form
    {
        private Button selectButton;
        private OpenFileDialog openFileDialog1;

        public string OpenFileDialogForm()
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


    }

    public class OpenFolder
    {
        private FolderBrowserDialog fbd;

        public string OpenFolderDialogForm()
        {
            fbd = new FolderBrowserDialog
            {
                Description = "Chose folder for Excel report."
            };

            // Show testDialog as a modal dialog
            DialogResult result = fbd.ShowDialog();
            string sfp = fbd.SelectedPath;

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                if (HasWritePermissionOnDir(sfp) == true)
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
        public bool HasWritePermissionOnDir(string path)
        {
            var writeAllow = false;
            var writeDeny = false;
            var accessControlList = Directory.GetAccessControl(path);
            if (accessControlList == null)
                return false;
            var accessRules = accessControlList.GetAccessRules(true, true, typeof(System.Security.Principal.SecurityIdentifier));
            if (accessRules == null)
                return false;

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

}