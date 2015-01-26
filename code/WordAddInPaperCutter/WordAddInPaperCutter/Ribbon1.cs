using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools;
using System.Windows.Forms;

namespace WordAddInPaperCutter
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            new EditorLogin().Show();
        }

        private void buttonQuestion_Click(object sender, RibbonControlEventArgs e)
        {
            if(Globals.ThisAddIn.editorID>0)
            {
                bool ISExit = false;
                foreach (CustomTaskPane ctp in Globals.ThisAddIn.CustomTaskPanes)
                {
                    if (ctp.Title.ToString().Equals("题目选择器"))
                    {
                        ctp.Visible = true;
                        ISExit = true;
                    }
                    else
                    {
                        ctp.Visible = false;
                    }
                }
                if (!ISExit)
                {
                    CustomTaskPane _customTaskPane = null;
                    UserControl1 u = new UserControl1();
                    _customTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(u, "题目选择器");
                    //_customTaskPane.Width = 1024;
                    _customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
                    _customTaskPane.Height = (int)(SystemInformation.WorkingArea.Height / 3);
                    _customTaskPane.Visible = true;
                }
            }
            else
            {
                new EditorLogin().Show();
            }
            
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            bool ISExit = false;
            foreach (CustomTaskPane ctp in Globals.ThisAddIn.CustomTaskPanes)
            {
                if (ctp.Title.ToString().Equals("截图工具"))
                {
                    ctp.Visible = true;
                    ISExit = true;
                }
                else
                {
                    ctp.Visible = false;
                }
            }
            if (!ISExit)
            {
                CustomTaskPane _customTaskPane = null;
                UserControlTest u = new UserControlTest();
                _customTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(u, "截图工具");
                //_customTaskPane.Width = 1024;
                _customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
                _customTaskPane.Height = (int)(SystemInformation.WorkingArea.Height / 3);
                _customTaskPane.Visible = true;
            }
        }
    }
}
