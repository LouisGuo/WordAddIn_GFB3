using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using WordAddInPaperCutter.Common;

namespace WordAddInPaperCutter
{
    public partial class EditorLogin : Form
    {
        private APIHelper apiHelper = new APIHelper();
        private JsonFileHelper jsonFileHelper = new Common.JsonFileHelper();
        private DB db = new DB();
        public EditorLogin()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string userName = this.textBoxUsername.Text.ToString();
            string password = this.textBoxPassword.Text.ToString();

            try
            {
                int returnUserID = apiHelper.CheckUser(userName, password);
                if (returnUserID > 0)
                {
                    this.Close();

                    Globals.ThisAddIn.editorID = returnUserID;

                    jsonFileHelper.WriteFileString(returnUserID + "", Globals.ThisAddIn.resourcesRootPath + "Configure\\user.txt");

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
                    MessageBox.Show("用户名或密码输入错误");
            }
            catch
            {
                MessageBox.Show("无法连接网络");

                string FileuserID = jsonFileHelper.GetFileString(Globals.ThisAddIn.resourcesRootPath + "Configure\\user.txt");
                if(!FileuserID.Equals(""))
                {
                    Globals.ThisAddIn.editorID = int.Parse(FileuserID);

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
            }
            
        }
    }
}
