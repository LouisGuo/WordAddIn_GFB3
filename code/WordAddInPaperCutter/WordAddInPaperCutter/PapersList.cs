using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordAddInPaperCutter.Common;
using WordAddInPaperCutter.JsonClass;
using Newtonsoft.Json;
using System.IO;
using Microsoft.Office.Tools;
using System.Net;

namespace WordAddInPaperCutter
{
    public partial class PapersList : Form
    {
        private DB db = new DB();
        private JsonFileHelper jsonFileHelper = new JsonFileHelper();
        private APIHelper apiHelper = new APIHelper();
        private const int BUTTON_TOP_POINT=20;
        private const int BUTTONHEIGHT = 39;
        private const int BUTTONWIDTH = 240;

        private int margintop = 30;
        private int exerciseOrderCount = 0;
        private int exerciseOrder = 0;

        private int paperNodeOrderCount = 0;
        private string paperName;

        private const int loadRightPaperWidth = 760;
        private double loadRightPaperZoom = 0.2;

        private string[] ChineseNumber = new string[] { "一", "二", "三", "四", "五", "六", "七", "八", "九", "十" };

        private List<int> deletedExerciseIndexList = new List<int>();
        public PapersList()
        {
            InitializeComponent();
            //设置模块和类型
            string model = jsonFileHelper.GetFileString(Globals.ThisAddIn.resourcesRootPath + "Configure\\model.json");
            string paperType = jsonFileHelper.GetFileString(Globals.ThisAddIn.resourcesRootPath + "Configure\\paperType.json");
            string problemType = jsonFileHelper.GetFileString(Globals.ThisAddIn.resourcesRootPath + "Configure\\problemType.json");

            List<CommonJson> tempJson = new List<CommonJson>();

            tempJson = (List<CommonJson>)JsonConvert.DeserializeObject(paperType, tempJson.GetType());

            if (tempJson != null)
            {
                for (int i = 0; i < tempJson.Count; i++)
                {
                    ListItem item = new ListItem(tempJson[i].Id.ToString(), tempJson[i].Name);
                    this.comboBoxPaperType.Items.Add(item);
                }
            }

            tempJson = new List<CommonJson>();
            tempJson = (List<CommonJson>)JsonConvert.DeserializeObject(problemType, tempJson.GetType());

            if (tempJson != null)
            {
                for (int i = 0; i < tempJson.Count; i++)
                {
                    ListItem item = new ListItem(tempJson[i].Id.ToString(), tempJson[i].Name);

                    this.comboBoxProblemSetType.Items.Add(item);
                }
            }
            


            this.splitContainer1.Height = SystemInformation.WorkingArea.Height;
            
            this.splitContainer1.BackColor = Color.Red;
            this.splitContainer1.Panel1.BackColor = Control.DefaultBackColor;
            this.splitContainer1.Panel2.BackColor = Color.White;

            //this.comboBox1.SelectedIndex = 0;

            //LoadLeftPapersList();

        }

        private void PapersList_Load(object sender, EventArgs e)
        {
            LoadLeftPapersList(0,"");
        }

        private void LoadLeftPapersList(int type,string keyWord)
        {
            this.lblSpeed.Visible = false;
            this.lblTime.Visible = false;
            this.progressBar1.Visible = false;

            this.panel2.Controls.Clear();
            this.panel1.Controls.Clear();
            this.panel3.Controls.Clear();

            this.groupBox1.Controls.Clear();

            this.panel7.Visible = true;
            this.groupBox1.Visible = true;
            if (this.label6.BackColor == Control.DefaultBackColor)
            {
                this.panel7.Visible = false;
                this.groupBox1.Visible = false;
            }
            
            string[] jsonName = Directory.GetFiles(Globals.ThisAddIn.exerciseJsonPath );

            int i = 0;
            int j = 0;
            foreach (string s in jsonName)
            {
                int index = s.LastIndexOf(".");
                string fileType = "";
                if (index + 5 == s.Length)
                    fileType = s.Substring(index, 5);

                if (fileType.ToLower().Equals(".json"))
                {
                    Paper thisTempPaper = new Paper();
                    ProblemSet thisTempProblemSet = new ProblemSet();
                    thisTempPaper = jsonFileHelper.GetPaperFromFile(s);
                    thisTempProblemSet = jsonFileHelper.GetProblemSetFromFile(s);


                    Button paper = new Button();
                    paper.Cursor = Cursors.Hand;
                    if (thisTempPaper.Name.ToString().Equals("") && thisTempProblemSet.Name.ToString().Equals(""))
                    {
                        continue;
                    }

                    //试卷名
                    if(type==0)
                    {
                        if(!keyWord.Equals(""))
                        {
                            //过滤掉不包含关键字的试卷
                            if (!thisTempPaper.Name.ToString().Contains(keyWord) && !thisTempProblemSet.Name.ToString().Contains(keyWord))
                                continue;
                        }
                    }
                    //试卷ID
                    else if(type==1)
                    {
                        if (!keyWord.Equals(""))
                        {
                            //过滤掉不包含关键字的试卷
                            if (!thisTempPaper.Id.ToString().Equals(keyWord) && !thisTempProblemSet.Id.ToString().Equals(keyWord))
                                continue;
                        }
                    }

                    if(thisTempPaper.PaperNodeList.Count!=0&&thisTempProblemSet.ExerciseList.Count==0)
                    {
                        paper.Text = thisTempPaper.Name.ToString();
                    }
                    if (thisTempPaper.PaperNodeList.Count == 0 && thisTempProblemSet.ExerciseList.Count != 0)
                    {
                        paper.Text = thisTempProblemSet.Name.ToString();
                    }
                    
                    string[] splitName = s.Split('\\');

                    string jsonFileName = splitName[splitName.Length - 1].Substring(0, splitName[splitName.Length - 1].Length - 5);
                    paper.Name = jsonFileName;
                    paper.TextAlign = ContentAlignment.MiddleLeft;
                    paper.Anchor = AnchorStyles.Right|AnchorStyles.Left | AnchorStyles.Top;
                    paper.Size = new System.Drawing.Size(BUTTONWIDTH, BUTTONHEIGHT);
                    paper.Click += new EventHandler(paperButton_Click);

                    CheckBox checkBox = new CheckBox();
                    checkBox.Text = "";
                    checkBox.Name = "CheckBox" + jsonFileName;
                    paper.Location = new Point(20, BUTTONHEIGHT * j + BUTTON_TOP_POINT);
                    checkBox.Location = new Point(5, BUTTONHEIGHT * j + BUTTON_TOP_POINT);

                    if(this.label1.BackColor==Control.DefaultBackColor)
                    {
                        if (jsonFileName.IndexOf("Uploaded") != 0)
                        {
                            this.groupBox1.Controls.Add(paper);
                            this.groupBox1.Controls.Add(checkBox);
                            j += 1;
                        }
                    }
                    else if(this.label2.BackColor==Control.DefaultBackColor)
                    {
                        if (jsonFileName.IndexOf("Uploaded") == 0)
                        {
                           
                            this.groupBox1.Controls.Add(paper);
                            this.groupBox1.Controls.Add(checkBox);
                            j += 1;
                        }
                    }
                    else if(this.label6.BackColor==Control.DefaultBackColor)
                    {

                    }

                }
            }


            this.groupBox1.Height = BUTTONHEIGHT * (j + 1) + BUTTON_TOP_POINT;
            if (BUTTONHEIGHT * (i + 1) + 102 < SystemInformation.WorkingArea.Height)
                this.splitContainer1.Height = SystemInformation.WorkingArea.Height;
            else
                this.splitContainer1.Height = BUTTONHEIGHT * (i + 1) + 102;


            if (BUTTONHEIGHT * (j + 1) + 102 < SystemInformation.WorkingArea.Height)
                this.splitContainer1.Height = SystemInformation.WorkingArea.Height;
            else
            {
                if (BUTTONHEIGHT * (j + 1) + 102 < BUTTONHEIGHT * (i + 1) + 102)
                    this.splitContainer1.Height = BUTTONHEIGHT * (i + 1) + 102;
                else
                    this.splitContainer1.Height = BUTTONHEIGHT * (j + 1) + 102;
            }


            bool HASBUTTON = false;
            foreach(Control c in this.groupBox1.Controls)
            {
                if(c is Button )
                {
                    Button button = c as Button;

                    paperName = button.Name.ToString();
                    HASBUTTON = true;
                    break;
                }
            }

            //if (!HASBUTTON)
            //    paperName = "exerciseUnClassified";

            if(HASBUTTON)
            {
                Button eventButton = new Button();
                eventButton.Name = paperName;
                paperButton_Click(eventButton, new EventArgs());
            }
            else
            {
                if(this.label6.Visible)
                {
                    paperName = "exerciseUnClassified";
                    Button eventButton = new Button();
                    eventButton.Name = paperName;
                    paperButton_Click(eventButton, new EventArgs());
                }
                this.label4.Visible = false;
            }

            
        }

        private void QuestionPictureBox_Click(object sender, EventArgs e)
        {
            PictureBox pb = sender as PictureBox;
            if(pb.BackColor!=Color.DarkGray)
            {
                foreach(Control c in this.panel1.Controls)
                {
                    if(c.Location.Y==pb.Location.Y)
                    {
                        c.BackColor = Color.DarkGray;
                    }
                }
                pb.BackColor = Color.DarkGray;
                int orderTemp = int.Parse(pb.Name.ToString().Substring(18, pb.Name.ToString().Length - 18));
                this.WindowState = FormWindowState.Minimized;

                bool ISExit = false;
                foreach (CustomTaskPane ctp in Globals.ThisAddIn.CustomTaskPanes)
                {
                    if (ctp.Title.ToString().Equals("题目详情编辑器"))
                    {
                        ctp.Visible = true;
                        ISExit = true;
                        UserControlEditExercise neww = ctp.Control as UserControlEditExercise;
                        neww.paperName = paperName;
                        neww.exerciseOrder = orderTemp;
                        neww.SetQuestionImg();
                    }
                    else
                    {
                        ctp.Visible = false;
                    }
                }
                if (!ISExit)
                {
                    CustomTaskPane _customTaskPane = null;
                    UserControlEditExercise u = new UserControlEditExercise(paperName, orderTemp);
                    _customTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(u, "题目详情编辑器");
                    //_customTaskPane.Width = 1024;
                    _customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
                    _customTaskPane.Height = (int)(SystemInformation.WorkingArea.Height / 3);
                    _customTaskPane.Visible = true;
                }
            }
        }

        private void PictureBox_MouseEnter(object sender, EventArgs e)
        {
            Control c = sender as Control;
            if (c.BackColor != Color.DarkGray)
            {
                foreach (Control con in this.panel1.Controls)
                {
                    if (con.Location.Y == c.Location.Y)
                    {
                        con.BackColor = Color.LightGray;
                    }
                } 
            }
            
            
        }

        private void PictureBox_MouseLeave(object sender, EventArgs e)
        {
            Control c = sender as Control;
            if (c.BackColor != Color.DarkGray)
            {
                foreach (Control con in this.panel1.Controls)
                {
                    if (con.Location.Y == c.Location.Y)
                    {
                        con.BackColor = Color.White;
                    }
                }
            }
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            Button b = sender as Button;
            foreach(Control c in this.panel1.Controls)
            {
                if (c is TextBox)
                    c.Enabled = false;
                if (c.Location.Y == b.Location.Y)
                    c.Visible = false;
            }

            int exerciseIndex = int.Parse(b.Name.ToString().Substring(12, b.Name.ToString().Length - 12));

            ProblemSet ps = jsonFileHelper.GetProblemSetFromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + ".json");
            Paper pa = jsonFileHelper.GetPaperFromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + ".json");

            deletedExerciseIndexList.Add(exerciseIndex);
            int oldExerciseIndex = exerciseIndex;
            for (int i = 0; i < deletedExerciseIndexList.Count; i++)
            {
                if (deletedExerciseIndexList[i] < oldExerciseIndex)
                    exerciseIndex -= 1;
            }

            if (ps.ExerciseList.Count!=0)
            {
                
                //ps.ExerciseList.RemoveAt(exerciseIndex);
                exerciseOrderCount = 0;
                exerciseOrder = exerciseIndex;
                DeleteExercise(ps.ExerciseList);
                jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(ps), Globals.ThisAddIn.exerciseJsonPath + paperName+".json");
               
            }
            else if(pa.PaperNodeList.Count!=0)
            {
                exerciseOrderCount = 0;
                exerciseOrder = exerciseIndex;
                DeleteExercise(pa.PaperNodeList);
                jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(pa), Globals.ThisAddIn.exerciseJsonPath + paperName + ".json");
               
            }
            else
            {
                MessageBox.Show("No ExerciseUnclassified");
            }
        }

        private void DeleteExercise(List<PaperNode> paperNodeList)
        {
            for(int i=0;i<paperNodeList.Count;i++)
            {
                exerciseOrderCount += 1;
                if(paperNodeList[i].ExerciseList.Count!=0)
                {
                    DeleteExercise(paperNodeList[i].ExerciseList);
                }
                if (paperNodeList[i].PaperNodeList.Count != 0)
                {
                    DeleteExercise(paperNodeList[i].PaperNodeList);
                }
            }

        }
        private void DeleteExercise(List<Exercise> exerciseList)
        {
            for (int i = 0; i < exerciseList.Count; i++)
            {
                if(exerciseOrderCount==exerciseOrder)
                {
                    exerciseList.RemoveAt(i);
                    return;
                }
                else
                {
                    exerciseOrderCount += 1;
                    if (exerciseList[i].ExerciseList.Count != 0)
                    {
                        DeleteExercise(exerciseList[i].ExerciseList);
                    }
                }
                
            }

        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Add");
        }

        private void TypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            this.splitContainer1.Focus();
            int type = comboBox.SelectedIndex;
            int exerciseIndex=int.Parse(comboBox.Name.ToString().Substring(8,comboBox.Name.ToString().Length-8));

            ProblemSet ps = jsonFileHelper.GetProblemSetFromFile(Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified.json");
            if (ps.ExerciseList.Count!=0)
            {
                int oldExerciseIndex = exerciseIndex;
                for (int i = 0; i < deletedExerciseIndexList.Count; i++)
                {
                    if (deletedExerciseIndexList[i] < oldExerciseIndex)
                        exerciseIndex -= 1;
                }
                ps.ExerciseList[exerciseIndex].QuestionType = type;

                jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(ps), Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified.json");
                
            }
            else
            {
                MessageBox.Show("No ExerciseUnclassified");
            }
        }

        private void LoadRightExercises(List<PaperNode> paperNodeList)
        {
            for(int i=0;i<paperNodeList.Count;i++)
            {
                PaperNode paperNodeTemp = paperNodeList[i] as PaperNode;

                Label label = new Label();

                label.Text = ChineseNumber[paperNodeOrderCount];



                label.Anchor = AnchorStyles.Left;
                label.Location = new Point(50, margintop);
                label.Size = new Size(40, 30);
                this.panel1.Controls.Add(label);
                
                PictureBox pictureBox = new PictureBox();
                pictureBox.Name = "PictureBoxQuestion" + exerciseOrderCount;
                pictureBox.Location = new Point(200, margintop);
                pictureBox.Anchor = AnchorStyles.Right | AnchorStyles.Left | AnchorStyles.Top;
                Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + paperNodeTemp.Title);
                loadRightPaperZoom = (double)loadRightPaperWidth / (double)(img.Width);


                pictureBox.Size = new Size((int)(img.Width * loadRightPaperZoom), (int)(img.Height * loadRightPaperZoom));
                pictureBox.Image = img;
                pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
                pictureBox.MouseEnter += new EventHandler(PictureBox_MouseEnter);
                pictureBox.MouseLeave += new EventHandler(PictureBox_MouseLeave);
                pictureBox.Click += new EventHandler(QuestionPictureBox_Click);
                this.panel1.Controls.Add(pictureBox);

                margintop += (int)(img.Height * loadRightPaperZoom);

                PictureBox pictureBoxLine = new PictureBox();
                //pictureBox1.Name = "PictureBoxQuestion" + exerciseOrderCount;
                pictureBoxLine.Location = new Point(30, margintop);
                pictureBoxLine.Anchor = AnchorStyles.Right | AnchorStyles.Left | AnchorStyles.Top;

                pictureBoxLine.Size = new Size(loadRightPaperWidth + 170, 1);
                pictureBoxLine.BackColor = Color.Gray;
                this.panel1.Controls.Add(pictureBoxLine);

                margintop += (int)(pictureBoxLine.Height);

                paperNodeOrderCount += 1;
                exerciseOrderCount += 1;
                

                if (paperNodeTemp.PaperNodeList.Count != 0)
                    LoadRightExercises(paperNodeTemp.PaperNodeList as List<PaperNode>);
                if (paperNodeTemp.ExerciseList.Count != 0)
                    LoadRightExercises(paperNodeTemp.ExerciseList as List<Exercise>);
            }
        }

        private void LoadRightExercises(List<Exercise> exerciseList)
        {
            
            for(int i=0;i<exerciseList.Count;i++)
            {
                Exercise exerciseTemp = exerciseList[i] as Exercise;

                if(this.exerciseUnClassified.BackColor==Color.Red)
                {
                    TextBox textBox = new TextBox();
                    textBox.Name = exerciseOrderCount + "";
                    //textBox.Text = exerciseOrderCount + "";

                    int text = 1;
                    if(exerciseTemp.QuestionType==1)
                    {
                        for(int j=i-1;j>=0;j--)
                        {
                            if (exerciseList[j].QuestionType == exerciseTemp.QuestionType)
                                text += 1;
                            else
                                break;
                        }
                        textBox.Text = "" + text;
                    }

                    char biaoti='A';
                    if (exerciseTemp.QuestionType == 0)
                    {
                        for (int j = i - 1; j >= 0; j--)
                        {
                            if (exerciseList[j].QuestionType == exerciseTemp.QuestionType)
                                biaoti += '\u0001';
                        }
                        textBox.Text = "" + biaoti;
                    }

                    if (exerciseTemp.QuestionType == 0)
                    {
                        textBox.Size = new Size(50, 30);
                        textBox.Location = new Point(30, margintop);
                    }
                    else if (exerciseTemp.QuestionType == 1)
                    {
                        textBox.Size = new Size(50, 30);
                        textBox.Location = new Point(30, margintop);
                    }
                    else if (exerciseTemp.QuestionType == 2)
                    {
                        textBox.Size = new Size(100, 30);
                        textBox.Location = new Point(30, margintop);
                    }
                    else if (exerciseTemp.QuestionType == 3)
                    {
                        textBox.Size = new Size(50, 30);
                        textBox.Location = new Point(80, margintop);
                    }
                    textBox.Anchor = AnchorStyles.Left;
                    textBox.MouseEnter += new EventHandler(PictureBox_MouseEnter);
                    textBox.MouseLeave += new EventHandler(PictureBox_MouseLeave);
                    this.panel1.Controls.Add(textBox);



                }
                else
                {
                    Label label = new Label();
                    label.Name = exerciseOrderCount + "";
                    int text = 1;
                    if (exerciseTemp.QuestionType == 1)
                    {
                        for (int j = i - 1; j >= 0; j--)
                        {
                            if (exerciseList[j].QuestionType == exerciseTemp.QuestionType)
                                text += 1;
                            else
                                break;
                        }
                        label.Text = "" + text;
                    }
                    label.Anchor = AnchorStyles.Left;
                    label.Location = new Point(50, margintop);
                    label.Size = new Size(40,30);
                    this.panel1.Controls.Add(label);

                    
                }

                ComboBox comboBox = new ComboBox();
                comboBox.Name = "ComboBox" + exerciseOrderCount;
                ListItem item1 = new ListItem("0", "标题");
                ListItem item2 = new ListItem("1", "小题");
                //ListItem item3 = new ListItem("2", "大题");
                //ListItem item4 = new ListItem("3", "大小");
                comboBox.Items.Add(item1);
                comboBox.Items.Add(item2);
                //comboBox.Items.Add(item3);
                //comboBox.Items.Add(item4);
                comboBox.Size = new Size(50, 30);
                comboBox.Location = new Point(90, margintop);
                comboBox.Anchor = AnchorStyles.Left;
                comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                if (this.exerciseUnClassified.BackColor != Color.Red)
                    comboBox.Enabled = false;
                comboBox.SelectedIndex = int.Parse(exerciseTemp.QuestionType.ToString());
                comboBox.MouseEnter += new EventHandler(PictureBox_MouseEnter);
                comboBox.MouseLeave += new EventHandler(PictureBox_MouseLeave);
                comboBox.SelectedIndexChanged += new EventHandler(TypeComboBox_SelectedIndexChanged);
                this.panel1.Controls.Add(comboBox);

                Button button = new Button();
                button.Cursor = Cursors.Hand;
                button.Name = "ButtonDelete" + exerciseOrderCount;
                button.BackColor = Control.DefaultBackColor;
                button.Text = "删除";
                button.Size = new Size(50, 22);
                button.Location = new Point(150, margintop);
                button.Anchor = AnchorStyles.Left;
                //button.Visible = false;

                button.Click += DeleteButton_Click;
                button.MouseEnter += new EventHandler(PictureBox_MouseEnter);
                button.MouseLeave += new EventHandler(PictureBox_MouseLeave);
                this.panel1.Controls.Add(button);

                //Button button1 = new Button();
                //button1.Cursor = Cursors.Hand;
                //button1.Name = "ButtonAdd" + exerciseOrderCount;
                //button1.BackColor = Control.DefaultBackColor;
                //button1.Text = "添加";
                //button1.Size = new Size(50, 22);
                //button1.Location = new Point(255, margintop);
                //button1.Anchor = AnchorStyles.Left;
                ////button1.Visible = false;

                //button1.Click += new EventHandler(AddButton_Click);
                //button1.MouseEnter += new EventHandler(PictureBox_MouseEnter);
                //button1.MouseLeave += new EventHandler(PictureBox_MouseLeave);
                //this.panel1.Controls.Add(button1);

                //Question
                PictureBox pictureBox = new PictureBox();
                pictureBox.Name = "PictureBoxQuestion" + exerciseOrderCount;
                pictureBox.Location = new Point(200, margintop);
                pictureBox.Anchor = AnchorStyles.Right | AnchorStyles.Left | AnchorStyles.Top;
                
                Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + exerciseTemp.Question);
                loadRightPaperZoom = (double)loadRightPaperWidth / (double)(img.Width);
                pictureBox.Size = new Size((int)(img.Width * loadRightPaperZoom), (int)(img.Height * loadRightPaperZoom));
                pictureBox.Image = img;
                pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
                pictureBox.MouseEnter += new EventHandler(PictureBox_MouseEnter);
                pictureBox.MouseLeave += new EventHandler(PictureBox_MouseLeave);
                pictureBox.Click += new EventHandler(QuestionPictureBox_Click);
                this.panel1.Controls.Add(pictureBox);

                margintop += (int)(img.Height * loadRightPaperZoom);
                if (!exerciseTemp.Answer.ToString().Equals("") && exerciseTemp.Answer.ToString().Contains(".png"))
                {
                    PictureBox pictureBox1 = new PictureBox();
                    //pictureBox1.Name = "PictureBoxQuestion" + exerciseOrderCount;
                    pictureBox1.Location = new Point(200, margintop);
                    pictureBox1.Anchor = AnchorStyles.Right | AnchorStyles.Left | AnchorStyles.Top;

                    Image img1 = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + exerciseTemp.Answer);
                    loadRightPaperZoom = (double)loadRightPaperWidth / (double)(img1.Width);
                    pictureBox1.Size = new Size((int)(img1.Width * loadRightPaperZoom), (int)(img1.Height * loadRightPaperZoom));
                    pictureBox1.Image = img1;
                    pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                    pictureBox1.MouseEnter += new EventHandler(PictureBox_MouseEnter);
                    pictureBox1.MouseLeave += new EventHandler(PictureBox_MouseLeave);
                    //pictureBox1.Click += new EventHandler(QuestionPictureBox_Click);
                    this.panel1.Controls.Add(pictureBox1);

                    margintop += (int)(img1.Height * loadRightPaperZoom);
                }

                if (!exerciseTemp.AnswerTips.ToString().Equals("") && exerciseTemp.AnswerTips.ToString().Contains(".png"))
                {
                    PictureBox pictureBox1 = new PictureBox();
                    //pictureBox1.Name = "PictureBoxQuestion" + exerciseOrderCount;
                    pictureBox1.Location = new Point(200, margintop);
                    pictureBox1.Anchor = AnchorStyles.Right | AnchorStyles.Left | AnchorStyles.Top;

                    Image img1 = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + exerciseTemp.AnswerTips);
                    loadRightPaperZoom = (double)loadRightPaperWidth / (double)(img1.Width);
                    pictureBox1.Size = new Size((int)(img1.Width * loadRightPaperZoom), (int)(img1.Height * loadRightPaperZoom));
                    pictureBox1.Image = img1;
                    pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                    pictureBox1.MouseEnter += new EventHandler(PictureBox_MouseEnter);
                    pictureBox1.MouseLeave += new EventHandler(PictureBox_MouseLeave);
                    //pictureBox1.Click += new EventHandler(QuestionPictureBox_Click);
                    this.panel1.Controls.Add(pictureBox1);

                    margintop += (int)(img1.Height * loadRightPaperZoom);
                }

                if (!exerciseTemp.Analysis.ToString().Equals("") && exerciseTemp.Analysis.ToString().Contains(".png"))
                {
                    PictureBox pictureBox1 = new PictureBox();
                    //pictureBox1.Name = "PictureBoxQuestion" + exerciseOrderCount;
                    pictureBox1.Location = new Point(200, margintop);
                    pictureBox1.Anchor = AnchorStyles.Right | AnchorStyles.Left | AnchorStyles.Top;

                    Image img1 = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + exerciseTemp.Analysis);
                    loadRightPaperZoom = (double)loadRightPaperWidth / (double)(img1.Width);
                    pictureBox1.Size = new Size((int)(img1.Width * loadRightPaperZoom), (int)(img1.Height * loadRightPaperZoom));
                    pictureBox1.Image = img1;
                    pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                    pictureBox1.MouseEnter += new EventHandler(PictureBox_MouseEnter);
                    pictureBox1.MouseLeave += new EventHandler(PictureBox_MouseLeave);
                    //pictureBox1.Click += new EventHandler(QuestionPictureBox_Click);
                    this.panel1.Controls.Add(pictureBox1);

                    margintop += (int)(img1.Height * loadRightPaperZoom);
                }



                PictureBox pictureBoxLine = new PictureBox();
                //pictureBox1.Name = "PictureBoxQuestion" + exerciseOrderCount;
                pictureBoxLine.Location = new Point(30, margintop);
                pictureBoxLine.Anchor = AnchorStyles.Right | AnchorStyles.Left | AnchorStyles.Top;

                pictureBoxLine.Size = new Size(loadRightPaperWidth+170, 1);
                pictureBoxLine.BackColor = Color.Gray;
                this.panel1.Controls.Add(pictureBoxLine);
                margintop += 1;

                margintop += (int)(pictureBoxLine.Height);
                exerciseOrderCount += 1;
                

                if (exerciseTemp.ExerciseList.Count != 0)
                    LoadRightExercises(exerciseTemp.ExerciseList as List<Exercise>);

            }
        }

        private void paperButton_Click(object sender, EventArgs e)
        {
            deletedExerciseIndexList.Clear();
            

            this.textBoxPaperName.Text = "";

            Button b = sender as Button;
            if (b.Name.ToString().Equals("exerciseUnClassified"))
                this.exerciseUnClassified.BackColor = Color.Red;
            else
                this.exerciseUnClassified.BackColor = Control.DefaultBackColor;
            foreach(Control c in this.groupBox1.Controls)
            {
                if(c is Button)
                {
                    Button button = c as Button;
                    if (button.Name == b.Name)
                        button.BackColor = Color.Red;
                    else
                        button.BackColor = Control.DefaultBackColor;
                }
            }
            paperName = b.Name.ToString();
            this.panel1.Controls.Clear();

            ProblemSet problemSet = jsonFileHelper.GetProblemSetFromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + ".json");
            Paper paper = jsonFileHelper.GetPaperFromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + ".json");

            

            if (paper.PaperNodeList.Count!=0&&problemSet.ExerciseList.Count==0)
            {
                exerciseOrderCount = 0;
                paperNodeOrderCount = 0;
                margintop = 30;
                LoadRightExercises(paper.PaperNodeList as List<PaperNode>);
            }
            else if (paper.PaperNodeList.Count==0&& problemSet.ExerciseList.Count!=0)
            {
                exerciseOrderCount = 0;
                paperNodeOrderCount = 0;
                margintop = 30;
                LoadRightExercises(problemSet.ExerciseList);
            }
            if (margintop + 70 > this.splitContainer1.Height)
                this.splitContainer1.Height = margintop + 100;
            //两次设置splitcontainer高度
            this.panel1.Controls.Clear();
            if (paper.PaperNodeList.Count != 0 && problemSet.ExerciseList.Count == 0)
            {
                exerciseOrderCount = 0;
                paperNodeOrderCount = 0;
                margintop = 30;
                LoadRightExercises(paper.PaperNodeList as List<PaperNode>);
                this.label4.Text ="试卷："+ paper.Name;
                this.label3.Text = "试卷总览";
            }
            else if (paper.PaperNodeList.Count == 0 && problemSet.ExerciseList.Count != 0)
            {
                exerciseOrderCount = 0;
                paperNodeOrderCount = 0;
                margintop = 30;
                LoadRightExercises(problemSet.ExerciseList as List<Exercise>);
                this.label4.Text = "习题集："+problemSet.Name;
                this.label3.Text = "习题集总览";
            }


            if (paperName.Equals("exerciseUnClassified"))
            {
                this.comboBoxPaperType.Visible = false;
                this.comboBoxProblemSetType.Visible = true;
                this.labelPaperName.Text = "习题集名：";
                for (int i = 0; i < problemSet.ExerciseList.Count;i++ )
                {
                    if (problemSet.ExerciseList[i].QuestionType == 0)
                    {
                        this.comboBoxPaperType.Visible = true;
                        this.comboBoxProblemSetType.Visible = false;
                        this.label4.Text = "试卷：" + paper.Name;
                        this.label3.Text = "试卷总览";
                        this.labelPaperName.Text = "试卷名：";
                        break;
                    }
                }

                this.label5.Visible = true;
                this.comboBoxPaperType.SelectedIndex = -1;
                this.comboBoxProblemSetType.SelectedIndex = -1;
                this.labelPaperName.Visible = true;
                this.textBoxPaperName.Visible = true;
                this.pictureBox3.Visible = true;
                this.pictureBox4.Visible = false;
                this.label4.Visible = false;
            }
            else
            {
                this.comboBoxPaperType.Visible = false;
                this.comboBoxProblemSetType.Visible = false;

                this.label4.Visible = true;
                this.label5.Visible = false;
                this.labelPaperName.Visible = false;
                this.textBoxPaperName.Visible = false;
                this.pictureBox4.Visible = true;
                this.pictureBox3.Visible = false;
            }

            this.pictureBox3.Location = new Point(this.pictureBox3.Location.X,this.panel1.Location.Y+this.panel1.Height+10);
            this.pictureBox4.Location = new Point(this.pictureBox4.Location.X, this.panel1.Location.Y + this.panel1.Height + 10);
            
        }

        private bool CheckEditorInput()
        {
            bool ISValid = true;
            if (textBoxPaperName.Text.ToString().Trim().Equals(""))
                ISValid = false;
            foreach (Control c in this.panel1.Controls)
            {
                if (c is TextBox)
                {
                    if ((c as TextBox).Text.ToString().Trim().Equals(""))
                    {
                        ISValid = false;
                    }
                }
            }

            return ISValid;
        }

       

        private void AddNewExerciseToProblemSet(List<Exercise> exerciseList,Exercise newExercise)
        {
            for (int i = 0; i <exerciseList.Count; i++)
            {
                if (exerciseOrder == exerciseOrderCount)
                {
                    exerciseList[i].ExerciseList.Add(newExercise);
                    return;
                }
                else
                {
                    exerciseOrderCount += 1;
                    if (exerciseList[i].ExerciseList.Count != 0)
                        AddNewExerciseToProblemSet(exerciseList[i].ExerciseList as List<Exercise>,newExercise);
                }
            }
        }

        private void AddNewExerciseToPaper(List<PaperNode> paperNodeList, Exercise newExercise)
        {
            for(int i=0;i<paperNodeList.Count;i++)
            {
                if (exerciseOrder == exerciseOrderCount)
                {
                    paperNodeList[i].ExerciseList.Add(newExercise);
                    return;
                }
                    
                else
                {
                    exerciseOrderCount += 1;
                    if (paperNodeList[i].PaperNodeList.Count != 0)
                        AddNewExerciseToPaper(paperNodeList[i].PaperNodeList as List<PaperNode>, newExercise);
                    if (paperNodeList[i].ExerciseList.Count != 0)
                        AddNewExerciseToProblemSet(paperNodeList[i].ExerciseList as List<Exercise>,newExercise);
                }
            }
        }

        private void AddNewPaperNodeToPaper(List<PaperNode> paperNodeList, Exercise newExercise)
        {
            for (int i = 0; i < paperNodeList.Count; i++)
            {
                if (exerciseOrder == exerciseOrderCount)
                {
                    PaperNode nodeTemp = new PaperNode();
                    nodeTemp.Title = newExercise.Question;
                    paperNodeList[i].PaperNodeList.Add(nodeTemp);
                    return;
                }
                    
                else
                {
                    exerciseOrderCount += 1;
                    if (paperNodeList[i].PaperNodeList.Count != 0)
                        AddNewPaperNodeToPaper(paperNodeList[i].PaperNodeList as List<PaperNode>, newExercise);
                    if (paperNodeList[i].ExerciseList.Count != 0)
                        AddNewExerciseToProblemSet(paperNodeList[i].ExerciseList as List<Exercise>, newExercise);
                }
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.editorID < 0)
            {
                MessageBox.Show("您还未登录，请先登录");
                new EditorLogin().Show();
                return;
            }

            if(!CheckEditorInput())
            {
                MessageBox.Show("输入框未填满");
                return;
            }
            else
            {
                this.ScrollControlIntoView(this.label3);
                paperName = this.textBoxPaperName.Text.ToString();

                string type = "problemSet";
                foreach (Control c in this.panel1.Controls)
                {
                    if (c is TextBox)
                    {
                        if ((c as TextBox).Text.ToString().Substring(0,1).Trim().ToString().ToLower().Equals("a"))
                        {
                            type = "paper";
                        }
                    }
                }

                ProblemSet exerciseUnClassified = jsonFileHelper.GetProblemSetFromFile(Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified.json");

                
                if(type.Equals("problemSet"))
                {
                    ProblemSet newProblemSet = new ProblemSet();
                    newProblemSet.Name = paperName;
                    newProblemSet.UploadUser = Globals.ThisAddIn.editorID;
                    foreach(Control c in this.panel1.Controls)
                    {
                        if (c is TextBox)
                        {
                            TextBox exerciseTextBox = c as TextBox;
                            string editorInput = exerciseTextBox.Text.ToString();
                            exerciseOrder = int.Parse(exerciseTextBox.Name.ToString());

                            exerciseOrderCount = 0;
                            //Exercise pointExercise = GetExerciseByOrder(exerciseUnClassified.exerciseList);
                            Exercise pointExercise = exerciseUnClassified.ExerciseList[exerciseOrder] as Exercise;
                            if(editorInput.IndexOf(".")<0)
                            {
                                
                                newProblemSet.ExerciseList.Add(pointExercise);
                            }
                            else
                            {
                                int lastDianIndex = editorInput.LastIndexOf(".");
                                string fatherInput = editorInput.Substring(0,lastDianIndex);
                                int fatherOrder = 0;

                                foreach(Control fatherC in this.panel1.Controls)
                                {
                                    if (fatherC is TextBox)
                                    {
                                        TextBox fatherTextBox = fatherC as TextBox;
                                        string fatherTextTemp = fatherTextBox.Text.ToString();
                                        int fatherOrderTemp = int.Parse(fatherTextBox.Name.ToString());
                                        if(fatherTextTemp.Equals(fatherInput)&&fatherOrderTemp<exerciseOrder)
                                        {
                                            fatherOrder = fatherOrderTemp;
                                            
                                        }
                                    }
                                }
                                exerciseOrder = fatherOrder;
                                exerciseOrderCount = 0;
                                AddNewExerciseToProblemSet(newProblemSet.ExerciseList as List<Exercise>, pointExercise);

                            }
                        }
                    }
                    if (!CheckDaTiLegality(newProblemSet.ExerciseList))
                    {
                        MessageBox.Show("输入错误，有大题后没有大题小题");
                        return;
                    }
                    //保存newProblemSet
                    DateTime now = DateTime.Now;
                    string paperNameTemp = now.Year + "" + now.Month + "" + now.Day + "" + now.Hour + "" + now.Minute + "" + now.Second + "" + new Random().Next(1000, 9999) + "____" + System.Guid.NewGuid().ToString();
                    
                    if(this.comboBoxProblemSetType.SelectedIndex>-1)
                    {
                        ListItem item = this.comboBoxProblemSetType.SelectedItem as ListItem;
                        newProblemSet.ProblemSetTypeId = int.Parse(item.ID);
                    }
                    
                    jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(newProblemSet), Globals.ThisAddIn.exerciseJsonPath + paperNameTemp + ".json");
                    jsonFileHelper.WriteFileString("", Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified.json");
                    
                    if(!Directory.Exists(Globals.ThisAddIn.exerciseJsonPath+paperNameTemp))
                    {
                        Directory.CreateDirectory(Globals.ThisAddIn.exerciseJsonPath + paperNameTemp);
                    }

                    this.panel1.Controls.Clear();
                    //for(int j=0;j<exerciseUnClassified.ExerciseList.Count;j++)
                    //{
                    //    string oldPath = Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified\\" + (exerciseUnClassified.ExerciseList[j] as Exercise).Question;
                    //    string newPath = Globals.ThisAddIn.exerciseJsonPath + paperNameTemp + "\\" + (exerciseUnClassified.ExerciseList[j] as Exercise).Question;
                    //    File.Copy(oldPath, newPath,true);
                    //}

                    string jsonResult = JsonConvert.SerializeObject(newProblemSet);
                    string[] imgName = Directory.GetFiles(Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified");
                    foreach (string s in imgName)
                    {
                        string[] nameSplit = s.Split('\\');
                        if (jsonResult.Contains(nameSplit[nameSplit.Length - 1]))
                        {
                            File.Copy(s, Globals.ThisAddIn.exerciseJsonPath + paperNameTemp + "\\" + nameSplit[nameSplit.Length - 1], true);
                        }
                    }

                }
                else if(type.Equals("paper"))
                {
                    Paper newPaper = new Paper();
                    newPaper.Name = paperName;
                    newPaper.UploadUser = Globals.ThisAddIn.editorID;

                    int lastPaperNodeOrder = 0;
                    foreach (Control c in this.panel1.Controls)
                    {
                        if (c is TextBox)
                        {
                            TextBox exerciseTextBox = c as TextBox;
                            string editorInput = exerciseTextBox.Text.ToString();
                            exerciseOrder = int.Parse(exerciseTextBox.Name.ToString());

                            exerciseOrderCount = 0;
                            //Exercise pointExercise = GetExerciseByOrder(exerciseUnClassified.exerciseList);
                            Exercise pointExercise = exerciseUnClassified.ExerciseList[exerciseOrder] as Exercise;
                            
                            //只有一位，没有father
                            if(editorInput.IndexOf(".")<0)
                            {
                                int output;
                                //题目1
                                if(int.TryParse(editorInput.Substring(0,1),out output))
                                {
                                    exerciseOrder = lastPaperNodeOrder;

                                    AddNewExerciseToPaper(newPaper.PaperNodeList as List<PaperNode>, pointExercise);
                                }
                                //paperNode 一
                                else
                                {
                                    lastPaperNodeOrder = exerciseOrder;
                                    PaperNode nodeTemp = new PaperNode();
                                    nodeTemp.Title = pointExercise.Question;
                                    newPaper.PaperNodeList.Add(nodeTemp);
                                }
                            }
                            //有father
                            else
                            {
                                int output;
                                //题目1.1
                                if (int.TryParse(editorInput.Substring(0, 1), out output))
                                {
                                    int lastDianIndex = editorInput.LastIndexOf(".");
                                    string fatherInput = editorInput.Substring(0, lastDianIndex);
                                    int fatherOrder = 0;

                                    foreach (Control fatherC in this.panel1.Controls)
                                    {
                                        if (fatherC is TextBox)
                                        {
                                            TextBox fatherTextBox = fatherC as TextBox;
                                            string fatherTextTemp = fatherTextBox.Text.ToString();
                                            int fatherOrderTemp = int.Parse(fatherTextBox.Name.ToString());
                                            if (fatherTextTemp.Equals(fatherInput) && fatherOrderTemp < exerciseOrder)
                                            {
                                                fatherOrder = fatherOrderTemp;

                                            }
                                        }
                                    }
                                    exerciseOrder = fatherOrder;
                                    exerciseOrderCount = 0;
                                    AddNewExerciseToPaper(newPaper.PaperNodeList as List<PaperNode>, pointExercise);
                                }
                                //paperNode 一.一
                                else
                                {
                                    int lastDianIndex = editorInput.LastIndexOf(".");
                                    string fatherInput = editorInput.Substring(0, lastDianIndex);
                                    int fatherOrder = 0;

                                    foreach (Control fatherC in this.panel1.Controls)
                                    {
                                        if (fatherC is TextBox)
                                        {
                                            TextBox fatherTextBox = fatherC as TextBox;
                                            string fatherTextTemp = fatherTextBox.Text.ToString();
                                            int fatherOrderTemp = int.Parse(fatherTextBox.Name.ToString());
                                            if (fatherTextTemp.Equals(fatherInput) && fatherOrderTemp < exerciseOrder)
                                            {
                                                fatherOrder = fatherOrderTemp;

                                            }
                                        }
                                    }
                                    lastPaperNodeOrder = exerciseOrder;
                                    exerciseOrder = fatherOrder;
                                    exerciseOrderCount = 0;

                                    AddNewPaperNodeToPaper(newPaper.PaperNodeList as List<PaperNode>, pointExercise);


                                }
                            }

                        }
                    }

                    if(!CheckDaTiLegality(newPaper.PaperNodeList))
                    {
                        MessageBox.Show("输入错误，有大题后没有大题小题，或试卷节点没有试题");
                        return;
                    }
                    //保存newProblemSet
                    DateTime now = DateTime.Now;

                    if (this.comboBoxPaperType.SelectedIndex > -1)
                    {
                        ListItem item = this.comboBoxPaperType.SelectedItem as ListItem;
                        newPaper.PaperTypeId = int.Parse(item.ID);
                    }

                    string paperNameTemp = now.Year + "" + now.Month + "" + now.Day + "" + now.Hour + "" + now.Minute + "" + now.Second + "" + new Random().Next(1000, 9999) + "____" + System.Guid.NewGuid().ToString();
                    jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(newPaper), Globals.ThisAddIn.exerciseJsonPath + paperNameTemp + ".json");
                    jsonFileHelper.WriteFileString("", Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified.json");
                    

                    if (!Directory.Exists(Globals.ThisAddIn.exerciseJsonPath + paperNameTemp))
                    {
                        Directory.CreateDirectory(Globals.ThisAddIn.exerciseJsonPath + paperNameTemp);
                    }

                    this.panel1.Controls.Clear();
                    //for (int j = 0; j < exerciseUnClassified.ExerciseList.Count; j++)
                    //{
                    //    string oldPath = Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified\\" + (exerciseUnClassified.ExerciseList[j] as Exercise).Question;
                    //    string newPath = Globals.ThisAddIn.exerciseJsonPath + paperNameTemp + "\\" + (exerciseUnClassified.ExerciseList[j] as Exercise).Question;
                    //    File.Copy(oldPath, newPath,true);
                    //}

                    string jsonResult = JsonConvert.SerializeObject(newPaper);
                    string[] imgName = Directory.GetFiles(Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified");
                    foreach (string s in imgName)
                    {
                        string[] nameSplit = s.Split('\\');
                        if (jsonResult.Contains(nameSplit[nameSplit.Length - 1]))
                        {
                            File.Copy(s, Globals.ThisAddIn.exerciseJsonPath + paperNameTemp + "\\" + nameSplit[nameSplit.Length - 1], true);
                        }
                    }
                }

            }
            MessageBox.Show("试卷已保存成功");
            LoadLeftPapersList(0,"");
        }

        private bool CheckDaTiLegality(List<PaperNode> paperNode)
        {
            for(int i=0;i<paperNode.Count;i++)
            {
                if (paperNode[i].PaperNodeList.Count == 0 && paperNode[i].ExerciseList.Count == 0)
                    return false;
                else
                {
                    if(paperNode[i].PaperNodeList.Count!=0)
                    {
                        bool result = CheckDaTiLegality(paperNode[i].PaperNodeList);
                        if (!result)
                            return false;
                    }
                    if(paperNode[i].ExerciseList.Count!=0)
                    {
                        bool result = CheckDaTiLegality(paperNode[i].ExerciseList);
                        if (!result)
                            return false;
                    }
                }
            }
            return true;
        }

        private bool CheckDaTiLegality(List<Exercise> exerciseList)
        {
            for (int i = 0; i < exerciseList.Count;i++ )
            {
                if (exerciseList[i].QuestionType == 2 && exerciseList[i].ExerciseList.Count == 0)
                    return false;
                else
                {
                    if(exerciseList[i].ExerciseList.Count!=0)
                    {
                        bool result = CheckDaTiLegality(exerciseList[i].ExerciseList);
                        if (!result)
                            return false;
                    }
                    
                }
            }
            return true;
        }

        private string UpLoadPaper(string thisPaperName,ProgressBar pb)
        {
            string result = "";
            

            Paper temPaper = jsonFileHelper.GetPaperFromFile(Globals.ThisAddIn.exerciseJsonPath + thisPaperName + ".json");
            ProblemSet tempProblemSet = jsonFileHelper.GetProblemSetFromFile(Globals.ThisAddIn.exerciseJsonPath  + thisPaperName + ".json");

            ZipClass zip = new ZipClass();
            if (Directory.Exists(Globals.ThisAddIn.exerciseJsonPath + "Temp"))
            {
                Directory.Delete(Globals.ThisAddIn.exerciseJsonPath + "Temp", true);
            }
            Directory.CreateDirectory(Globals.ThisAddIn.exerciseJsonPath + "Temp");
            string jsonResult = jsonFileHelper.GetFileString(Globals.ThisAddIn.exerciseJsonPath + thisPaperName + ".json");
            int returnPaperID = 0;

            if (temPaper.PaperNodeList.Count != 0 && tempProblemSet.ExerciseList.Count == 0)
            {
                returnPaperID = apiHelper.sendPaperJson_Request(jsonResult.Replace("\"", "'"));
            }
            else if (temPaper.PaperNodeList.Count == 0 && tempProblemSet.ExerciseList.Count != 0)
            {
                returnPaperID = apiHelper.sendProblemSetJson_Request(jsonResult.Replace("\"", "'"));
            }
            

            Directory.CreateDirectory(Globals.ThisAddIn.exerciseJsonPath + "Temp\\" + thisPaperName);
            Directory.CreateDirectory(Globals.ThisAddIn.exerciseJsonPath + "Uploaded" + thisPaperName);

            string[] imgName = Directory.GetFiles(Globals.ThisAddIn.exerciseJsonPath + thisPaperName);
            foreach (string s in imgName)
            {
                string[] nameSplit = s.Split('\\');
                if (jsonResult.Contains(nameSplit[nameSplit.Length - 1]))
                {
                    File.Copy(s, Globals.ThisAddIn.exerciseJsonPath + "Temp\\" + thisPaperName + "\\" + nameSplit[nameSplit.Length - 1], true);
                    File.Copy(s, Globals.ThisAddIn.exerciseJsonPath + "Uploaded" + thisPaperName + "\\" + nameSplit[nameSplit.Length - 1], true);
                }

            }

            //把上传后的试卷放到Uploaded文件夹下
            if (temPaper.PaperNodeList.Count != 0 && tempProblemSet.ExerciseList.Count == 0)
            {
                temPaper.Id = returnPaperID;
                jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(temPaper), Globals.ThisAddIn.exerciseJsonPath + thisPaperName + ".json");
            }
            else if (temPaper.PaperNodeList.Count == 0 && tempProblemSet.ExerciseList.Count != 0)
            {
                tempProblemSet.Id = returnPaperID;
                jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(tempProblemSet), Globals.ThisAddIn.exerciseJsonPath + thisPaperName + ".json");
            }

            File.Move(Globals.ThisAddIn.exerciseJsonPath + thisPaperName + ".json", Globals.ThisAddIn.exerciseJsonPath + "Uploaded" + thisPaperName + ".json");
            zip.ZipFileFromDirectory(Globals.ThisAddIn.exerciseJsonPath + "Temp", Globals.ThisAddIn.exerciseJsonPath + thisPaperName + ".zip", 0);
            

            if (temPaper.PaperNodeList.Count != 0 && tempProblemSet.ExerciseList.Count == 0)
            {
                result = apiHelper.UploadPaper_Request(returnPaperID, Globals.ThisAddIn.exerciseJsonPath + thisPaperName + ".zip", thisPaperName + "", this.progressBar1, this.lblTime, this.lblSpeed);
            }
            else if (temPaper.PaperNodeList.Count == 0 && tempProblemSet.ExerciseList.Count != 0)
            {
                result = apiHelper.UploadProblemSet_Request(returnPaperID, Globals.ThisAddIn.exerciseJsonPath + thisPaperName + ".zip", thisPaperName + "", pb, this.lblTime, this.lblSpeed);
            }

            
            return result;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                this.ScrollControlIntoView(this.label3);
                this.progressBar1.Visible = true;
                this.lblTime.Visible = true;
                this.lblSpeed.Visible = true;

                string result = UpLoadPaper(paperName,this.progressBar1);

                this.progressBar1.Visible = false;
                this.lblTime.Visible = false;
                this.lblSpeed.Visible = false;
                MessageBox.Show("上传成功:" + result);

                if(!result.Equals("1"))
                {
                    string sss = result;
                }
            }
            catch
            {
                MessageBox.Show("无法连接到服务器");
            }


            LoadLeftPapersList(0,"");
        }


        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            foreach(Control c in this.panel2.Controls)
            {
                if(c is CheckBox)
                {
                    CheckBox checkBox = c as CheckBox;
                    checkBox.Checked = checkBox1.Checked;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(Globals.ThisAddIn.exerciseJsonPath + "Deleted"))
            {
                Directory.CreateDirectory(Globals.ThisAddIn.exerciseJsonPath + "Deleted");
            }
            Control.ControlCollection collection = this.panel2.Controls;

            if(this.tabControl1.SelectedTab==this.tabPage1)
            {
                collection = this.panel2.Controls;
            }
            else if(this.tabControl1.SelectedTab==this.tabPage2)
            {
                collection = this.panel3.Controls;
            }
            int fileDeleteCount = 0;
            foreach(Control c in collection)
            {
                if(c is CheckBox)
                {
                    CheckBox checkBox = c as CheckBox;
                    if(checkBox.Checked)
                    {
                        string jsonFileName = checkBox.Name.ToString().Substring(8, checkBox.Name.ToString().Length-8);
                        Directory.CreateDirectory(Globals.ThisAddIn.exerciseJsonPath + "Deleted\\" + jsonFileName);
                        File.Move(Globals.ThisAddIn.exerciseJsonPath + jsonFileName + ".json", Globals.ThisAddIn.exerciseJsonPath + "Deleted\\" + jsonFileName + ".json");

                        string[] imgName = Directory.GetFiles(Globals.ThisAddIn.exerciseJsonPath + jsonFileName);
                        foreach (string s in imgName)
                        {
                            string[] nameSplit = s.Split('\\');
                            File.Copy(s, Globals.ThisAddIn.exerciseJsonPath + "Deleted\\" + jsonFileName + "\\" + nameSplit[nameSplit.Length - 1],true);
                            
                        }
                        fileDeleteCount += 1;
                    }
                }
            }
            MessageBox.Show("一共删除"+fileDeleteCount+"份试卷");
            LoadLeftPapersList(0,"");
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Control c in this.panel3.Controls)
            {
                if (c is CheckBox)
                {
                    CheckBox checkBox = c as CheckBox;
                    checkBox.Checked = checkBox2.Checked;
                }
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Label label = sender as Label;
            label.BackColor = Control.DefaultBackColor;

            foreach(Control c in this.panel4.Controls)
            {
                if(c.Name!=label.Name)
                {
                    c.BackColor = Color.DarkGray;
                }
            }

            ListItem item1 = new ListItem("0", "按试卷名查找");
            ListItem item2 = new ListItem("1", "按试卷ID查找");
            this.comboBox1.Items.Clear();
            if (this.label1.BackColor == Control.DefaultBackColor)
            {
                this.comboBox1.Items.Add(item1);
                this.comboBox1.SelectedIndex = 0;
            }
            else if (this.label2.BackColor == Control.DefaultBackColor)
            {
                this.comboBox1.Items.Add(item1);
                this.comboBox1.Items.Add(item2);
                this.comboBox1.SelectedIndex = 0;
            }
            else if (this.label6.BackColor == Control.DefaultBackColor)
            {

            }
            LoadLeftPapersList(0,"");
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Control c in this.groupBox1.Controls)
            {
                if (c is CheckBox)
                {
                    CheckBox checkBox = c as CheckBox;
                    checkBox.Checked = checkBox3.Checked;
                }
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(Globals.ThisAddIn.exerciseJsonPath + "Deleted"))
            {
                Directory.CreateDirectory(Globals.ThisAddIn.exerciseJsonPath + "Deleted");
            }
            int fileDeleteCount = 0;
            foreach (Control c in this.groupBox1.Controls)
            {
                if (c is CheckBox)
                {
                    CheckBox checkBox = c as CheckBox;
                    if (checkBox.Checked)
                    {
                        string jsonFileName = checkBox.Name.ToString().Substring(8, checkBox.Name.ToString().Length - 8);
                        Directory.CreateDirectory(Globals.ThisAddIn.exerciseJsonPath + "Deleted\\" + jsonFileName);
                        File.Move(Globals.ThisAddIn.exerciseJsonPath + jsonFileName + ".json", Globals.ThisAddIn.exerciseJsonPath + "Deleted\\" + jsonFileName + ".json");

                        string[] imgName = Directory.GetFiles(Globals.ThisAddIn.exerciseJsonPath + jsonFileName);
                        foreach (string s in imgName)
                        {
                            string[] nameSplit = s.Split('\\');
                            File.Copy(s, Globals.ThisAddIn.exerciseJsonPath + "Deleted\\" + jsonFileName + "\\" + nameSplit[nameSplit.Length - 1],true);

                        }
                        fileDeleteCount += 1;
                    }
                }
            }
            MessageBox.Show("一共删除" + fileDeleteCount + "份试卷");
            LoadLeftPapersList(0,"");
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            
            try
            {
                int count = 0;
                string result = "";
                foreach (Control c in this.groupBox1.Controls)
                {
                    if (c is CheckBox)
                    {
                        CheckBox cb = c as CheckBox;
                        if (cb.Checked)
                        {
                            foreach (Control c1 in this.groupBox1.Controls)
                            {
                                if (c1 is Button)
                                {
                                    Button b = c1 as Button;
                                    if (b.Location.Y == cb.Location.Y)
                                    {
                                        ProgressBar pb = new ProgressBar();
                                        pb.Size = b.Size;
                                        pb.Location = b.Location;
                                        this.groupBox1.Controls.Remove(b);
                                        this.groupBox1.Controls.Add(pb);
                                        pb.BringToFront();
                                        result += UpLoadPaper(b.Name.ToString(), pb);

                                        count += 1;

                                    }
                                }
                            }
                        }
                    }
                }

                MessageBox.Show("一共上传了" + count + "份试卷 result:" + result);
            }
            catch
            {
                MessageBox.Show("无法连接到服务器");
            }
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            LoadLeftPapersList(this.comboBox1.SelectedIndex,this.textBox1.Text.ToString().Trim());
            
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            if(this.textBox1.ForeColor==Color.LightGray)
            {
                this.textBox1.Text = "";
                this.textBox1.ForeColor = Color.Black;
            }
        }
    }
}
