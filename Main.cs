using KeePass.Plugins;
using exclel;
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

namespace ExportXlsAll
{
    public class ExportXlsAllExt : Plugin
    {
        public static IPluginHost m_host = null;
        public override bool Initialize(IPluginHost host)  // main init method
        {
            if (host == null) return false;
            m_host = host;
            return true;
        }

        public override void Terminate()
        {

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
                exportGroup.Click += exclel.Exporter.ExportGroup;
                container.DropDownItems.Add(exportGroup);

                exportAll.Text = "Export all to .xls";
                exportAll.Click += exclel.Exporter.ExportDatabase;
                container.DropDownItems.Add(exportAll);

                exportSelected.Text = "Export selected to .xls";
                exportSelected.Click += exclel.Exporter.ExportSelected;
                container.DropDownItems.Add(exportSelected);
                return container;

                
            }
            return null;
        }
    }
}