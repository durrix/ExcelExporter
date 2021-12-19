using KeePass.Plugins;
using KeePassLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportXlsAll
{
    public class ExportXlsAllExt : Plugin
    {
        private IPluginHost m_host = null;

        //Public fields

        public static bool is_chosen = false;    //selection form communication fields
        public static List<string> keys = new List<string>();
        public static List<string> selectedKeys = new List<string>();
        public string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";


        public override bool Initialize(IPluginHost host)  // main init method
        {
            if (host == null) return false;
            m_host = host;

            return true;
        }

        public override void Terminate()
        {

        }
        private void ExportDatabase(object sender, EventArgs e)
        {


            var exelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = exelapp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

            int lastUsedColumn = 5; //
            Dictionary<string, int> columnNames = new Dictionary<string, int>(); //holds indexes of coumn names already in use
            columnNames.Add("Title", 1);
            columnNames.Add("UserName", 2);
            columnNames.Add("URL", 3);
            columnNames.Add("Password", 4);
            columnNames.Add("Notes", 5);

            worksheet.Cells[1, 1] = "Title";   //main column names
            worksheet.Cells[1, 1].Font.Bold = true;
            worksheet.Cells[1, 2] = "UserName";
            worksheet.Cells[1, 2].Font.Bold = true;
            worksheet.Cells[1, 3] = "URL";
            worksheet.Cells[1, 3].Font.Bold = true;
            worksheet.Cells[1, 4] = "Password";
            worksheet.Cells[1, 4].Font.Bold = true;
            worksheet.Cells[1, 5] = "Notes";
            worksheet.Cells[1, 5].Font.Bold = true;

            for (int i = 1; i <= 5; i++)
            {
                worksheet.Cells[1, i].Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                worksheet.Cells[1, i].ColumnWidth = 25;
                worksheet.Cells[1, i].Font.Bold = true;
            }


            var workSheet_range = worksheet.get_Range("A1", "F4");//adding Background Color
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();//adding Background Color
            worksheet.get_Range("A1", "H1").Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//Align to center

            var elements = m_host.Database.RootGroup.GetEntries(false).ToArray();

            for (int i = 0; i < elements.Length; i++)              // for each element in database
            {
                var keys = elements[i].Strings.GetKeys().ToArray(); // get keys for current element
                for (int j = 0; j < keys.Length; j++)
                {
                    bool exists = false;
                    var existingNames = columnNames.Keys.ToArray(); //
                    foreach (string columnName in existingNames)    // check if current column name is already in use
                    {
                        if (columnName == keys[j])
                        {
                            exists = true;
                        }
                    }

                    if (exists)
                    {
                        worksheet.Cells[i + 2, columnNames[keys[j]]] = elements[i].Strings.ReadSafe(keys[j]);  //if column is already exists write down under it
                    }
                    else
                    {
                        workSheet_range = worksheet.get_Range("A1", alphabet[lastUsedColumn] + (elements.Length + 1).ToString());//Creating Border
                        workSheet_range.Borders.Color = Color.Black.ToArgb();//Creating Border

                        worksheet.Cells[1, lastUsedColumn + 1].Interior.Color = ColorTranslator.ToOle(Color.LightGreen);//adding Background Color
                        worksheet.Cells[1, lastUsedColumn + 1].ColumnWidth = 25; //Changing Size

                        workSheet_range = worksheet.get_Range("A1", alphabet[lastUsedColumn] + (elements.Length + 1).ToString());//Creating Border
                        workSheet_range.Borders.Color = Color.Black.ToArgb();//Creating Border

                        worksheet.Cells[1, lastUsedColumn + 1] = keys[j];                                       //if column does not exist, create it and add to existing columns
                        columnNames.Add(keys[j], lastUsedColumn + 1);
                        worksheet.Cells[i + 2, lastUsedColumn + 1] = elements[i].Strings.ReadSafe(keys[j]);
                        lastUsedColumn++;
                    }
                }
            }
            var sfd = new SaveFileDialog();
            sfd.Filter = "Excel (*.xls)|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(sfd.FileName);       //save
            }
            exelapp.Quit();
            var processes = Process.GetProcessesByName("Microsoft Excel");
            foreach (Process proc in processes)
            {
                proc.Kill();
            }
        }
        async private void ExportSelected(object sender, EventArgs e)
        {


            var exelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = exelapp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

            int lastUsedColumn = 5;
            Dictionary<string, int> columnNames = new Dictionary<string, int>();
            columnNames.Add("Title", 1);
            columnNames.Add("UserName", 2);
            columnNames.Add("URL", 3);
            columnNames.Add("Password", 4);
            columnNames.Add("Notes", 5);

            worksheet.Cells[1, 1] = "Title"; //main column names
            worksheet.Cells[1, 1].Font.Bold = true;
            worksheet.Cells[1, 2] = "UserName";
            worksheet.Cells[1, 2].Font.Bold = true;
            worksheet.Cells[1, 3] = "URL";
            worksheet.Cells[1, 3].Font.Bold = true;
            worksheet.Cells[1, 4] = "Password";
            worksheet.Cells[1, 4].Font.Bold = true;
            worksheet.Cells[1, 5] = "Notes";
            worksheet.Cells[1, 5].Font.Bold = true;
            for (int i = 1; i <= 5; i++)
            {
                worksheet.Cells[1, i].Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                worksheet.Cells[1, i].ColumnWidth = 25;
                worksheet.Cells[1, i].Font.Bold = true;
            }

            if (m_host.MainWindow.GetSelectedEntriesCount() == 0)
            {
                MessageBox.Show("No entries selected");
                goto Release;
            }
            else
            {
                var elements = m_host.MainWindow.GetSelectedEntries().ToArray(); //get selected element


                for (int i = 0; i < elements.Length; i++)
                {
                    keys = elements[i].Strings.GetKeys(); //get 

                    selectedKeys.Clear();

                    selectedKeys.Add("UserName");
                    selectedKeys.Add("URL");
                    selectedKeys.Add("Password");
                    selectedKeys.Add("Title");
                    selectedKeys.Add("Notes");

                    var workSheet_range = worksheet.get_Range("A1", "F2");//adding Background Color
                    workSheet_range.Borders.Color = Color.Black.ToArgb();//adding Background Color

                    SelectionForm selectionForm = new SelectionForm(); //create selection form instance
                    selectionForm.Show(); //open selection form 

                    while (!is_chosen) { await Task.Delay(100); } //wait until ok is pressed
                    is_chosen = false;

                    var sKeys = selectedKeys.ToArray(); //get selected advanced keys
                    for (int j = 0; j < sKeys.Length; j++)
                    {
                        bool exists = false; //existance flag
                        var existingNames = columnNames.Keys.ToArray();
                        foreach (string columnName in existingNames) // check if column is already exists 
                        {
                            if (columnName == sKeys[j])
                            {
                                exists = true;
                            }
                        }

                        if (exists)
                        {
                            worksheet.Cells[i + 2, columnNames[sKeys[j]]] = elements[i].Strings.ReadSafe(sKeys[j]); //if exists write down under the column
                        }
                        else
                        {
                            workSheet_range = worksheet.get_Range("A1", alphabet[lastUsedColumn] + (elements.Length + 1).ToString());//Creating Border
                            workSheet_range.Borders.Color = Color.Black.ToArgb();//Creating Border

                            worksheet.Cells[1, lastUsedColumn + 1].Interior.Color = ColorTranslator.ToOle(Color.LightGreen);//adding Background Color
                            worksheet.Cells[1, lastUsedColumn + 1].ColumnWidth = 25; //Changing Size

                            workSheet_range = worksheet.get_Range("A1", alphabet[lastUsedColumn] + (elements.Length + 1).ToString());//Creating Border
                            workSheet_range.Borders.Color = Color.Black.ToArgb();//Creating Border

                            worksheet.Cells[1, lastUsedColumn + 1] = sKeys[j];                                      // if does not exists, create new column and add to used names dictionary
                            columnNames.Add(sKeys[j], lastUsedColumn + 1);
                            worksheet.Cells[i + 2, lastUsedColumn + 1] = elements[i].Strings.ReadSafe(sKeys[j]);
                            worksheet.Cells[1, lastUsedColumn + 1].Font.Bold = true;
                            lastUsedColumn++;
                        }
                    }
                }
            }


            var sfd = new SaveFileDialog();
            sfd.Filter = "Excel (*.xls)|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK) //if sfd returns success
            {
                workbook.SaveAs(sfd.FileName); //save as excel file
            }
        Release:
            exelapp.Quit();
            var processes = Process.GetProcessesByName("Microsoft Excel");
            foreach (Process proc in processes)
            {
                proc.Kill();
            }

        }


        async private void ExportGroup(object sender, EventArgs e)
        {

            var exelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = exelapp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

            worksheet.Cells[1, 1] = "Title"; //main column names
            worksheet.Cells[1, 1].Font.Bold = true;
            worksheet.Cells[1, 2] = "UserName";
            worksheet.Cells[1, 2].Font.Bold = true;
            var workSheet_range = worksheet.get_Range("A1", "B2");//adding Background Color
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();//adding Background Color
            worksheet.get_Range("A1", "B2").Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//Align to center
            for (int i = 1; i <= 2; i++)
            {
                worksheet.Cells[1, i].Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                worksheet.Cells[1, i].ColumnWidth = 25;
                worksheet.Cells[1, i].Font.Bold = true;
            }
            int lastUsedColumn = 2;
            Dictionary<string, int> columnNames = new Dictionary<string, int>();
            columnNames.Add("Title", 1);
            columnNames.Add("UserName", 2);

            var sGroup = m_host.MainWindow.GetSelectedGroup();

            if (sGroup == null)
                goto Exit;

            var subGroupsNames = new List<string>();
            //creating a list of subgroup names
            foreach (var subGroup in sGroup.GetGroups(false).ToArray())
            {
                subGroupsNames.Add(subGroup.Name);
            }


            var entries = sGroup.GetEntries(true).ToArray(); //contains subgroup elements

            for (int i = 0; i < entries.Length; i++)
            {
                if (EntryContainsGroupName(subGroupsNames.ToArray(), entries[i].Strings.ReadSafe("Title")))
                {


                    keys.Clear();
                    selectedKeys.Clear();
                    keys = entries[i].Strings.GetKeys();
                    SelectionForm form = new SelectionForm();
                    form.Show();
                    while (!is_chosen) { await Task.Delay(100); }
                    is_chosen = false;

                    selectedKeys.Add("Title");
                    selectedKeys.Add("UserName");
                    var sKeys = selectedKeys.ToArray();

                    for (int j = 0; j < sKeys.Length; j++)
                    {
                        bool exists = false;
                        var cNames = columnNames.Keys.ToArray();
                        for (int k = 0; k < cNames.Length; k++)
                        {
                            if (cNames[k] == sKeys[j])
                            {
                                exists = true;
                                break;
                            }
                        }

                        if (exists)
                        {
                            worksheet.Cells[i + 2, columnNames[sKeys[j]]] = entries[i].Strings.ReadSafe(sKeys[j]);
                        }
                        else
                        {
                            workSheet_range = worksheet.get_Range("A1", alphabet[lastUsedColumn] + (entries.Length + 1).ToString());//Creating Border
                            workSheet_range.Borders.Color = Color.Black.ToArgb();//Creating Border

                            worksheet.Cells[1, lastUsedColumn + 1].Interior.Color = ColorTranslator.ToOle(Color.LightGreen);//adding Background Color
                            worksheet.Cells[1, lastUsedColumn + 1].ColumnWidth = 25; //Changing Size

                            worksheet.Cells[1, lastUsedColumn + 1] = sKeys[j];                                      // if does not exists, create new column and add to used names dictionary
                            columnNames.Add(sKeys[j], lastUsedColumn + 1);
                            worksheet.Cells[i + 2, lastUsedColumn + 1] = entries[i].Strings.ReadSafe(sKeys[j]);
                            worksheet.Cells[1, lastUsedColumn + 1].Font.Bold = true;
                            lastUsedColumn++;
                        }
                    }

                }
            }
            var sfd = new SaveFileDialog();
            sfd.Filter = "Excel (*.xls)|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK) //if sfd returns success        
                workbook.SaveAs(sfd.FileName); //save as excel file
            Exit:
            exelapp.Quit();
            var processes = Process.GetProcessesByName("Microsoft Excel");
            foreach (Process proc in processes)
            {
                proc.Kill();
            }
        }



        private bool EntryContainsGroupName(string[] names, string entryName)
        {
            for (int i = 0; i < names.Length; i++)
            {
                if (names[i] == entryName)
                {
                    return true;
                }
            }
            return false;
        }


        public override ToolStripMenuItem GetMenuItem(PluginMenuType t) //add menu controls 
        {
            ToolStripMenuItem container = new ToolStripMenuItem("Export to .xls");
            ToolStripMenuItem exportSelected = new ToolStripMenuItem();
            ToolStripMenuItem exportAll = new ToolStripMenuItem();
            ToolStripMenuItem exportGroup = new ToolStripMenuItem();
            if (t == PluginMenuType.Main)
            {
                exportGroup.Text = "Export subgroups to .xls";
                exportGroup.Click += this.ExportGroup;
                container.DropDownItems.Add(exportGroup);

                exportAll.Text = "Export all to .xls";
                exportAll.Click += this.ExportDatabase;
                container.DropDownItems.Add(exportAll);

                exportSelected.Text = "Export selected to .xls";
                exportSelected.Click += this.ExportSelected;
                container.DropDownItems.Add(exportSelected);
                return container;
            }
            return null;
        }
    }
}