using exclel;
using KeePassLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
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
        // constants used to hide a checkbox
        public const int TVIF_STATE = 0x8;
        public const int TVIS_STATEIMAGEMASK = 0xF000;
        public const int TV_FIRST = 0x1100;
        public const int TVM_SETITEM = TV_FIRST + 63;

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam,
        IntPtr lParam);

        // struct used to set node properties
        public struct TVITEM
        {
            public int mask;
            public IntPtr hItem;
            public int state;
            public int stateMask;
            [MarshalAs(UnmanagedType.LPTStr)]
            public String lpszText;
            public int cchTextMax;
            public int iImage;
            public int iSelectedImage;
            public int cChildren;
            public IntPtr lParam;

        }
        private void HideCheckBox(TreeNode node)
        {
            TVITEM tvi = new TVITEM();
            tvi.hItem = node.Handle;
            tvi.mask = TVIF_STATE;
            tvi.stateMask = TVIS_STATEIMAGEMASK;
            tvi.state = 0;
            IntPtr lparam = Marshal.AllocHGlobal(Marshal.SizeOf(tvi));
            Marshal.StructureToPtr(tvi, lparam, false);
            SendMessage(node.TreeView.Handle, TVM_SETITEM, IntPtr.Zero, lparam);
        }
        void tree_DrawNode(object sender, DrawTreeNodeEventArgs e)
        {
            if (e.Node.Level == 0)
            {
                HideCheckBox(e.Node);
                e.DrawDefault = true;
            }
            else
            {
                e.Graphics.DrawString(e.Node.Text, e.Node.TreeView.Font,
                   Brushes.Black, e.Node.Bounds.X, e.Node.Bounds.Y);
            }
        }
        private void SelectionForm_Load(object sender, EventArgs e)
        {

            var elements = exclel.Exporter.elems.ToArray();
            treeView.DrawMode = TreeViewDrawMode.OwnerDrawText;
            treeView.DrawNode += new DrawTreeNodeEventHandler(tree_DrawNode);

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
                    root.Checked = false;
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