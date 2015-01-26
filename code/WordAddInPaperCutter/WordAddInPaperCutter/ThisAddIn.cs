using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;
using WordAddInPaperCutter.Common;
using WordAddInPaperCutter.JsonClass;
using System.IO;
using Newtonsoft.Json;
using System.Windows.Forms;
using System.Threading;

namespace WordAddInPaperCutter
{
    public partial class ThisAddIn
    {

        public CustomTaskPane _customTaskPane = null;
        private JsonFileHelper jsonFileHelper = new JsonFileHelper();
        private APIHelper apiHelper = new APIHelper();
        private ProblemSet _exerciseUnClassified = new ProblemSet();
        private string _exerciseJsonPath = @"C:\Resources\Papers\";
        private string _resourcesRootPath = @"C:\Resources\";

        private int _editorID = -10;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            UserControl1 u = new UserControl1();
            _customTaskPane = this.CustomTaskPanes.Add(u, "题目选择器");

            //_customTaskPane.Width = 1024;
            _customTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionBottom;
            _customTaskPane.Height = (int)(SystemInformation.WorkingArea.Height / 3);
            _customTaskPane.Visible = true;

            CleanDash();

            Thread t = new Thread(UpdateConfig);
            t.Start();
            
        }

        private void UpdateConfig()
        {
            

            //每次程序开始时，获取模块知识点，模块，试卷类型，习题集类型
            try
            {
                if (!Directory.Exists(resourcesRootPath + "Configure"))
                {
                    Directory.CreateDirectory(resourcesRootPath + "Configure");
                }

                string paperTypeJson = apiHelper.GetPaperType();
                if (!File.Exists(resourcesRootPath + "Configure\\paperType.json"))
                {
                    File.Create(resourcesRootPath + "Configure\\paperType.json");
                }
                if (!paperTypeJson.Equals(""))
                    jsonFileHelper.WriteFileString(paperTypeJson, resourcesRootPath + "Configure\\paperType.json");

                string problemTypeJson = apiHelper.GetProblemType();
                if (!File.Exists(resourcesRootPath + "Configure\\problemType.json"))
                {
                    File.Create(resourcesRootPath + "Configure\\problemType.json");
                }
                if (!problemTypeJson.Equals(""))
                    jsonFileHelper.WriteFileString(problemTypeJson, resourcesRootPath + "Configure\\problemType.json");

                string knowlege = apiHelper.GetKnowlege(1);
                if (!File.Exists(resourcesRootPath + "Configure\\knowlege.json"))
                {
                    File.Create(resourcesRootPath + "Configure\\knowlege.json");
                }
                if (!knowlege.Equals(""))
                    jsonFileHelper.WriteFileString(knowlege, resourcesRootPath + "Configure\\knowlege.json");

                string model = apiHelper.GetModel(1);
                if (!File.Exists(resourcesRootPath + "Configure\\model.json"))
                {
                    File.Create(resourcesRootPath + "Configure\\model.json");
                }
                if (!model.Equals(""))
                    jsonFileHelper.WriteFileString(model, resourcesRootPath + "Configure\\model.json");
            }
            catch
            {
                MessageBox.Show("无法连接到服务器，无法更新知识树");
            }

        }

        private void CleanDash()
        {
            try
            {
                if (!Directory.Exists(_exerciseJsonPath))
                {
                    Directory.CreateDirectory(_exerciseJsonPath);
                }
                if (!Directory.Exists(_exerciseJsonPath + "Deleted"))
                {
                    Directory.CreateDirectory(_exerciseJsonPath + "Deleted");
                }
                if (!Directory.Exists(_exerciseJsonPath + "Uploaded"))
                {
                    Directory.CreateDirectory(_exerciseJsonPath + "Uploaded");
                }
                if (!Directory.Exists(_exerciseJsonPath + "exerciseUnClassified"))
                {
                    Directory.CreateDirectory(_exerciseJsonPath + "exerciseUnClassified");
                }

                //每次程序开始时，清理已组卷的题目图片

                string[] zipName = Directory.GetFiles(Globals.ThisAddIn.exerciseJsonPath);

                foreach (string s in zipName)
                {
                    int index = s.LastIndexOf(".");
                    string fileType = "";
                    if (index + 4 == s.Length)
                        fileType = s.Substring(index, 4);

                    if (fileType.ToLower().Equals(".zip"))
                    {

                        File.Delete(s);
                    }
                }

                //每次程序开始时，清理已组卷的题目图片
                if (!Directory.Exists(_exerciseJsonPath))
                {
                    Directory.CreateDirectory(_exerciseJsonPath);
                }

                string jsonStr = jsonFileHelper.GetFileString(_exerciseJsonPath + "exerciseUnClassified.json");
                ProblemSet exerciseUnClassified = new ProblemSet();
                exerciseUnClassified = jsonFileHelper.GetProblemSetFromFile(_exerciseJsonPath + "exerciseUnClassified.json");
                string[] allImg = Directory.GetFiles(Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified");
                if (exerciseUnClassified.ExerciseList.Count != 0)
                {
                    foreach (string s in allImg)
                    {
                        string[] nameSplit = s.Split('\\');
                        if (!jsonStr.Contains(nameSplit[nameSplit.Length - 1]))
                            File.Delete(s);
                    }
                }
                else
                {
                    foreach (string s in allImg)
                    {
                        File.Delete(s);
                    }
                }

                //每次程序开始时，清理已上传的试卷图片
                string[] imgDire = Directory.GetDirectories(Globals.ThisAddIn.exerciseJsonPath);
                foreach (string d in imgDire)
                {
                    string[] dSplit = d.Split('\\');
                    string jsonFileName = dSplit[dSplit.Length - 1];
                    string[] jsonFile = Directory.GetFiles(Globals.ThisAddIn.exerciseJsonPath);
                    bool ISEX = false;
                    foreach (string s in jsonFile)
                    {
                        string[] jsonFileSplit = s.Split('\\');
                        string fullJsonFileName = jsonFileSplit[jsonFileSplit.Length - 1];
                        if ((jsonFileName + ".json").Equals(fullJsonFileName))
                            ISEX = true;
                    }

                    if (!ISEX && (!jsonFileName.Equals("Uploaded") && !jsonFileName.Equals("Papers") && !jsonFileName.Equals("Deleted")))
                        Directory.Delete(Globals.ThisAddIn.exerciseJsonPath + jsonFileName, true);
                }
            }
            catch
            {
                MessageBox.Show("无法清理垃圾文件");
            }
            
        }

        public int editorID
        {

            get { return _editorID; }
            set { _editorID = value; }
        }

        public string exerciseJsonPath
        {
            get { return _exerciseJsonPath; }
            set { _exerciseJsonPath = value; }
        }

        public string resourcesRootPath
        {
            get { return _resourcesRootPath; }
            set { _resourcesRootPath = value; }
        }

        public ProblemSet exerciseUnClassified
        {
            get { return _exerciseUnClassified; }
            set { _exerciseUnClassified = value; }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
