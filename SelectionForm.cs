using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportXlsAll
{
    public partial class SelectionForm : Form
    {
        public SelectionForm()
        {
            InitializeComponent();
        }

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            var keys = ExportXlsAllExt.keys.ToArray();
            foreach(string key in keys)
            {
                if(key!="UserName" && key!= "URL" && key!= "Notes" && key != "Title" && key != "Password") //export advanced items to CheckListBox
                ListBox.Items.Add(key, false);
            }
            
        }

        private void ok_btn_Click(object sender, EventArgs e)
        {
            var items = ListBox.CheckedItems;
            
            foreach(var item in items) //add selected advanced strings to list
            {
                ExportXlsAllExt.selectedKeys.Add(item.ToString());
            }
            ExportXlsAllExt.is_chosen = true; 
            this.Close(); //close form
        }
    }
}
