using KeePass.Plugins;
using ExportXlsAll;
using KeePassLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace exclel
{
    class Exporter
    {

        //imports
        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        private static extern int FindWindow(string sClass, string sWindow);

        // glob fields
        public static bool is_chosen = false;    //selection form communication fields
        public static List<PwEntry> elems = new List<PwEntry>();
        public static List<PwEntry> elements1 = new List<PwEntry>();
        public static List<string> subGroupsNames = new List<string>();
        public static Dictionary<PwEntry, List<string>> g_pairs = new Dictionary<PwEntry, List<string>>();
        public static string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public static bool isMatched = false;


        public static void ExportDatabase(object sender, EventArgs e)
        {
            var exelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = exelapp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;


            string savedir = "";
            //save directory
            var sfd = new SaveFileDialog();
            sfd.Filter = "Excel (*.xls)|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
                savedir = sfd.FileName;
            else
                goto Exit;








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

            var elements = ExportXlsAllExt.m_host.Database.RootGroup.GetEntries(false).ToArray();

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

        TryAgain:
            try
            {
                workbook.SaveAs(savedir);
            }
            catch (Exception ioex)
            {
                MessageBox.Show("Close Exel", "Eroor"); // err message
                goto TryAgain;
            }

        Exit:
            exelapp.Quit();
        }


        async static public void ExportSelected(object sender, EventArgs e)
        {

            var exelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = exelapp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

            string savedir = "";

            var sfd = new SaveFileDialog();
            sfd.Filter = "Excel (*.xls)|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK) //if sfd returns success
            {
                savedir = sfd.FileName;
            }
            else goto Release;

            int lastUsedColumn = 5;
            var processes = Process.GetProcessesByName("Microsoft Excel");

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

            if (ExportXlsAll.ExportXlsAllExt.m_host.MainWindow.GetSelectedEntriesCount() == 0)
            {
                MessageBox.Show("No entries selected");
                goto Release;
            }
            else
            {
                var elements = ExportXlsAll.ExportXlsAllExt.m_host.MainWindow.GetSelectedEntries(); //get selected element
                elems.Clear();
                elems = new List<PwEntry>(elements);
                g_pairs.Clear();
                SelectionForm selectionForm = new SelectionForm(); //create selection form instance
                selectionForm.Show();
                while (!is_chosen) { await Task.Delay(100); } //wait until ok is pressed
                is_chosen = false;

                for (int i = 0; i < g_pairs.Keys.Count; i++)
                {
                    g_pairs[g_pairs.Keys.ElementAt(i)].Add("UserName");
                    g_pairs[g_pairs.Keys.ElementAt(i)].Add("URL");
                    g_pairs[g_pairs.Keys.ElementAt(i)].Add("Password");
                    g_pairs[g_pairs.Keys.ElementAt(i)].Add("Title");
                    g_pairs[g_pairs.Keys.ElementAt(i)].Add("Notes");

                    var workSheet_range = worksheet.get_Range("A1", "F2");//adding Background Color
                    workSheet_range.Borders.Color = Color.Black.ToArgb();//adding Background Color




                    var sKeys = g_pairs[g_pairs.Keys.ElementAt(i)].ToArray(); //get selected advanced keys
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
                            worksheet.Cells[i + 2, columnNames[sKeys[j]]] = g_pairs.Keys.ElementAt(i).Strings.ReadSafe(sKeys[j]); //if exists write down under the column
                        }
                        else
                        {
                            workSheet_range = worksheet.get_Range("A1", alphabet[lastUsedColumn] + (g_pairs.Keys.Count + 1).ToString());//Creating Border
                            workSheet_range.Borders.Color = Color.Black.ToArgb();//Creating Border

                            worksheet.Cells[1, lastUsedColumn + 1].Interior.Color = ColorTranslator.ToOle(Color.LightGreen);//adding Background Color
                            worksheet.Cells[1, lastUsedColumn + 1].ColumnWidth = 25; //Changing Size

                            workSheet_range = worksheet.get_Range("A1", alphabet[lastUsedColumn] + (g_pairs.Keys.Count + 1).ToString());//Creating Border
                            workSheet_range.Borders.Color = Color.Black.ToArgb();//Creating Border

                            worksheet.Cells[1, lastUsedColumn + 1] = sKeys[j];                                      // if does not exists, create new column and add to used names dictionary
                            columnNames.Add(sKeys[j], lastUsedColumn + 1);
                            worksheet.Cells[i + 2, lastUsedColumn + 1] = g_pairs.Keys.ElementAt(i).Strings.ReadSafe(sKeys[j]);
                            worksheet.Cells[1, lastUsedColumn + 1].Font.Bold = true;
                            lastUsedColumn++;
                        }
                    }
                }
            }
        
        Tryagain:
            try
            {
                workbook.SaveAs(savedir);
            }
            catch (Exception ioex)
            {
                MessageBox.Show("Close Excel", "Error"); // err message
                goto Tryagain;
            }

        Release:
            exelapp.Quit();

        }



        async static public void ExportGroup(object sender, EventArgs e)
        {

            var exelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = exelapp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

            string savedir = "";

            var sfd = new SaveFileDialog();
            sfd.Filter = "Excel (*.xls)|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK) //if sfd returns success
            {
                savedir = sfd.FileName;
            }
            else goto Release;

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

            var sGroup = ExportXlsAllExt.m_host.MainWindow.GetSelectedGroup();

            if (sGroup == null)
                goto Release;

            
            //creating a list of subgroup names
            subGroupsNames.Clear();
            foreach (var subGroup in sGroup.GetGroups(false).ToArray())
            {
                subGroupsNames.Add(subGroup.Name);

            }
            

            elems.Clear();
            g_pairs.Clear();
            elems = sGroup.GetEntries(true).ToList(); //contains subgroup elements

            elements1.Clear();

            foreach (var name in Exporter.subGroupsNames)
            {
                foreach (var elem in elems)
                {
                    if (name == elem.Strings.ReadSafe("Title"))
                    {
                        elements1.Add(elem);
                    }
                }
            }
            elems = elements1;

            SelectionForm form = new SelectionForm();
            form.Show();
            while (!is_chosen) { await Task.Delay(100); }
            is_chosen = false;

            for (int i = 0; i < g_pairs.Keys.Count; i++)
            {
                if (util.Util.EntryContainsGroupName(subGroupsNames.ToArray(), g_pairs.Keys.ElementAt(i).Strings.ReadSafe("Title")))
                {

                    g_pairs[g_pairs.Keys.ElementAt(i)].Add("Title");
                    g_pairs[g_pairs.Keys.ElementAt(i)].Add("UserName");
                    var sKeys = g_pairs[g_pairs.Keys.ElementAt(i)].ToArray();

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
                            worksheet.Cells[i + 2, columnNames[sKeys[j]]] = g_pairs.Keys.ElementAt(i).Strings.ReadSafe(sKeys[j]);
                        }
                        else
                        {
                            workSheet_range = worksheet.get_Range("A1", alphabet[lastUsedColumn] + (g_pairs.Keys.Count + 1).ToString());//Creating Border
                            workSheet_range.Borders.Color = Color.Black.ToArgb();//Creating Border

                            worksheet.Cells[1, lastUsedColumn + 1].Interior.Color = ColorTranslator.ToOle(Color.LightGreen);//adding Background Color
                            worksheet.Cells[1, lastUsedColumn + 1].ColumnWidth = 25; //Changing Size

                            worksheet.Cells[1, lastUsedColumn + 1] = sKeys[j];                                      // if does not exists, create new column and add to used names dictionary
                            columnNames.Add(sKeys[j], lastUsedColumn + 1);
                            worksheet.Cells[i + 2, lastUsedColumn + 1] = g_pairs.Keys.ElementAt(i).Strings.ReadSafe(sKeys[j]);
                            worksheet.Cells[1, lastUsedColumn + 1].Font.Bold = true;
                            lastUsedColumn++;
                        }
                    }

                }
            }
        Tryagain:
            try
            {
                workbook.SaveAs(savedir);
            }
            catch (Exception ioex)
            {
                MessageBox.Show("Close Excel", "Error"); // err message
                goto Tryagain;
            }

        Release:
            exelapp.Quit();
        }
    }
}



namespace word
{
    public class Exporter
    {



    }
}




namespace util
{
    public class Util
    {
        public static bool EntryContainsGroupName(string[] names, string entryName)
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
    }

}

