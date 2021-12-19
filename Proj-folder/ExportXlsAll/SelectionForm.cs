using exclel;
using KeePassLib;
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
            var elements = exclel.Exporter.elems.ToArray();

            foreach (PwEntry element in elements)
            {
                var keys = element.Strings.GetKeys().ToArray();
                try
                {
                    TreeNode root = new TreeNode(element.Strings.ReadSafe("Title"));
                    treeView.Nodes.Add(root);
                    foreach (string key in keys)
                    {
                        if (element.Strings.ReadSafe(key) != "" && key != "UserName" && key != "URL" && key != "Title" && key != "Password" && key != "Notes")
                        {
                            TreeNode branch = new TreeNode(key);
                            root.Nodes.Add(branch);
                        }
                    }
                }
                catch
                {

                }

            }

        }

        private void ok_btn_Click(object sender, EventArgs e)
        {
            Dictionary<PwEntry, List<string>> entrykeys = new Dictionary<PwEntry, List<string>>();
            var elements = exclel.Exporter.elems.ToArray();
            foreach (PwEntry element in elements)
            {
                List<string> sKeysList = new List<string>();
                foreach (TreeNode root in treeView.Nodes)
                {
                    if (root.Text == element.Strings.ReadSafe("Title"))
                    {
                        foreach (TreeNode branch in root.Nodes)
                        {
                            if (branch.Checked)
                            {
                                sKeysList.Add(branch.Text);
                            }
                        }
                    }
                }
                Exporter.g_pairs = entrykeys;
                entrykeys.Add(element, sKeysList);
                Exporter.is_chosen = true;
                this.Close();
            }


        }
    }
}