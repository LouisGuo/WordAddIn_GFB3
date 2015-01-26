using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordAddInPaperCutter.Common;
using WordAddInPaperCutter.JsonClass;
using System.IO;
using Newtonsoft.Json;
using System.Drawing.Imaging;

namespace WordAddInPaperCutter
{
    public partial class UserControlEditExercise : UserControl
    {
        public string paperName="";
        public int exerciseOrder = 0;
        private int exerciseOrderCount = 0;
        private ProblemSet problemSet = new ProblemSet();
        private JsonFileHelper jsonFileHelper = new JsonFileHelper();
        private ImgHelper imgHelper = new ImgHelper();
        private Paper paper = new Paper();

        private Exercise pointExercise = new Exercise();

        private int groupBoxQuestion_Height;
        private int groupBoxTips_Height;
        private int groupBoxAnswer2_Height;
        private int groupBoxAnalysis_Height;

        private int groupBoxQuestion_Location;
        private int groupBoxTips_Location;
        private int groupBoxAnswer_Location;
        private int groupBoxAnalysis_Location;

        private int pictureBoxQuestion_Height;
        private int pictureBoxTips_Height;
        private int pictureBoxAnswer_Height;
        private int pictureBoxAnalysis_Height;


        private int panel_Height;

        public UserControlEditExercise(string paperName,int exerciseOrder)
        {
            InitializeComponent();

            LoadKonwlegeTree("");
            this.panel3.Parent = this.panel1;
            this.panel4.Parent = this.panel1;
            this.panel5.Parent = this.panel1;
            this.panel6.Parent = this.panel1;

            this.panel3.Location = new Point(this.panel3.Location.X, 3);
            this.panel4.Location = new Point(this.panel4.Location.X, 3);
            this.panel5.Location = new Point(this.panel5.Location.X, 3);
            this.panel6.Location = new Point(this.panel6.Location.X, 3);

            SetTabVisible();

            groupBoxQuestion_Location = this.groupBoxQuestion.Location.Y;
            groupBoxAnswer_Location = this.groupBoxAnswer2.Location.Y;
            groupBoxTips_Location = this.groupBoxTips.Location.Y;
            groupBoxAnalysis_Location = this.groupBoxAnalysis.Location.Y;

            groupBoxQuestion_Height = this.groupBoxQuestion.Height;
            groupBoxTips_Height = this.groupBoxTips.Height;
            groupBoxAnswer2_Height = this.groupBoxAnswer2.Height;
            groupBoxAnalysis_Height = this.groupBoxAnalysis.Height;

            pictureBoxQuestion_Height=this.pictureBoxQuestion.Height;
            pictureBoxTips_Height = this.pictureBoxTips.Height;
            pictureBoxAnswer_Height = this.pictureBoxAnswer.Height;
            pictureBoxAnalysis_Height = this.pictureBoxAnalysis.Height;
            panel_Height = this.panel3.Height;

            this.paperName = paperName;
            this.exerciseOrder = exerciseOrder;
            SetQuestionImg();
            //this.pictureBoxQuestion.Image = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath+paperName+"\\"+pictureName);
        }

        private void LoadKonwlegeTree(string keyWord)
        {
            this.treeView1.Nodes.Clear();
            string knowlegeList = jsonFileHelper.GetFileString(Globals.ThisAddIn.resourcesRootPath + "Configure\\knowlege.json");

            List<Knowlege> tempJson = new List<Knowlege>();
            tempJson = JsonConvert.DeserializeObject<List<Knowlege>>(knowlegeList);


            for (int i = 0; i < tempJson.Count; i++)
            {
                if(!keyWord.Equals("")&&!tempJson[i].Name.ToString().Contains(keyWord))
                {
                    continue;
                }
                TreeNode node = new TreeNode();
                node.Name = tempJson[i].Id.ToString();
                node.Text = tempJson[i].Name.ToString();

                Knowlege knowlege = tempJson[i] as Knowlege;

                this.treeView1.Nodes.Add(node);
                //if (!AddTreeNode(this.treeView1, knowlege))
                //    this.treeView1.Nodes.Add(node);

            }
        }

        private bool AddTreeNode(TreeView tree, Knowlege knowlege)
        {
            foreach(TreeNode node in tree.Nodes)
            {
                if(node.Name.ToString().Equals(knowlege.FatherId.ToString()))
                {
                    TreeNode temp = new TreeNode();
                    temp.Name = knowlege.Id.ToString();
                    temp.Text = knowlege.Name;
                    node.Nodes.Add(temp);
                    return true;
                }
                if(node.Nodes.Count!=0)
                {
                    foreach(TreeNode nodeSon in node.Nodes)
                    {
                        bool result = AddTreeNode(nodeSon.TreeView, knowlege);
                        if (result)
                            return true;
                    }
                }
            }
            return false;
        }

        private void SetGroupBoxLocation()
        {
            this.groupBoxQuestion.Location = new Point(this.groupBoxQuestion.Location.X, groupBoxQuestion_Location);
            this.groupBoxTips.Location = new Point(this.groupBoxTips.Location.X, groupBoxTips_Location);
            this.groupBoxAnswer2.Location = new Point(this.groupBoxAnswer2.Location.X, groupBoxAnswer_Location);
            this.groupBoxAnalysis.Location = new Point(this.groupBoxAnalysis.Location.X, groupBoxAnalysis_Location);

            this.groupBoxQuestion.Height = groupBoxQuestion_Height;
            this.groupBoxTips.Height = groupBoxTips_Height;
            this.groupBoxAnswer2.Height = groupBoxAnswer2_Height;
            this.groupBoxAnalysis.Height = groupBoxAnalysis_Height;

            this.pictureBoxQuestion.Height = pictureBoxQuestion_Height;
            this.pictureBoxTips.Height = pictureBoxTips_Height;
            this.pictureBoxAnswer.Height = pictureBoxAnswer_Height;
            this.pictureBoxAnalysis.Height = pictureBoxAnalysis_Height;

            this.pictureBoxAnalysis.Image = null;
            this.pictureBoxAnswer.Image = null;
            this.pictureBoxQuestion.Image = null;
            this.pictureBoxTips.Image = null;
            //this.panel1.Height = panel1_Height;

            this.panel3.Height = panel_Height;
            this.panel4.Height = panel_Height;
            this.panel5.Height = panel_Height;
            this.panel6.Height = panel_Height;

            this.pictureBox3.Location = new Point(this.pictureBox3.Location.X, this.groupBoxQuestion.Location.Y + this.groupBoxQuestion.Height + 10);
            this.pictureBox5.Location = new Point(this.pictureBox5.Location.X, this.groupBoxAnswer2.Location.Y + this.groupBoxAnswer2.Height + 10);
            this.pictureBox6.Location = new Point(this.pictureBox6.Location.X, this.groupBoxAnalysis.Location.Y + this.groupBoxAnalysis.Height + 10);
            this.pictureBox7.Location = new Point(this.pictureBox7.Location.X, this.groupBoxTips.Location.Y + this.groupBoxTips.Height + 10);

            this.textBox2.Text = "";


        }

        public void SetQuestionImg()
        {
            
            string filePath = Globals.ThisAddIn.exerciseJsonPath + paperName+".json";
            paper = jsonFileHelper.GetPaperFromFile(filePath);
            problemSet = jsonFileHelper.GetProblemSetFromFile(filePath);

            //Exercise pointExercise = new Exercise();
            if (paper.PaperNodeList.Count != 0 && problemSet.ExerciseList.Count == 0)
            {
                exerciseOrderCount = 0;
                pointExercise = GetExerciseByOrder(paper.PaperNodeList as List<PaperNode>);
            }
            else if (paper.PaperNodeList.Count == 0 && problemSet.ExerciseList.Count != 0)
            {
                exerciseOrderCount = 0;
                pointExercise = GetExerciseByOrder(problemSet.ExerciseList as List<Exercise>);
            }

            ReSetQuestionDetail();
        }

        private void ReSetQuestionDetail()
        {
            SetGroupBoxLocation();


            this.listView1.Items.Clear();
            RemoveAllSelectedNodeTab(this.treeView1);
            
            if (pointExercise.Question != null)
            {

                int marginTopAdd = 0;
                double zoom = 0.2;

                if(pointExercise.QuestionType>0)
                {
                    this.label10.Visible = true;
                    this.label12.Visible = true;
                    this.label13.Visible = true;
                    //this.groupBoxAnswer.Visible = true;
                    //this.groupBoxAnswer2.Visible = true;
                    this.pictureBox4.Visible = true;
                    this.pictureBox5.Visible = true;

                    this.label4.Visible = true;
                    this.comboBoxType.Visible = true;
                }
                else if (pointExercise.QuestionType == 0)
                {
                    this.label10.Visible = false;
                    this.label12.Visible = false;
                    this.label13.Visible = false;
                    this.groupBoxAnswer.Visible = false;
                    this.groupBoxAnswer2.Visible = false;
                    this.pictureBox4.Visible = false;
                    this.pictureBox5.Visible = false;

                    this.label4.Visible = false;
                    this.label5.Visible = false;
                    this.comboBox1.Visible = false;
                    this.comboBoxType.Visible = false;
                }

                if(pointExercise.Type>0)
                    this.comboBoxType.SelectedIndex=pointExercise.Type-1;
                if (pointExercise.AnswerNumber>0)
                {
                    this.comboBox1.SelectedIndex = pointExercise.AnswerNumber-1;
                }
                if (pointExercise.Score>0)
                {
                    this.textBox4.Text = pointExercise.Score.ToString();
                }
                if (pointExercise.PredictDifficult>0)
                {
                    this.textBox3.Text = pointExercise.PredictDifficult.ToString();
                }
                else
                {
                    //this.textBox3.Text = "";
                }
                if (pointExercise.Split>0)
                {
                    this.textBox5.Text=pointExercise.Split.ToString();
                }
                else
                {
                    //this.textBox5.Text = "";
                }

                if (!pointExercise.Source.ToString().Equals(""))
                {
                    this.textBox1.Text = pointExercise.Source.ToString();
                    this.textBox1.ForeColor = Color.Black;
                }
                else
                {
                    
                    //this.textBox1.Text = "2013-北京-高考-N（输入格式）";
                    //this.textBox1.ForeColor = Color.LightGray;
                }
                if (!pointExercise.Video.ToString().Equals(""))
                {
                    this.textBox2.Text = pointExercise.Video.ToString();
                }
                else
                {

                }
                if (pointExercise.PredictDifficult>0)
                {
                    this.textBox3.Text = pointExercise.PredictDifficult.ToString();
                }


                if (!pointExercise.Question.ToString().Equals(""))
                {
                    Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + pointExercise.Question);
                    zoom = (double)this.pictureBoxQuestion.Width / (double)img.Width;
                    marginTopAdd = ((int)(img.Height * zoom) - this.pictureBoxQuestion.Height);
                    //this.pictureBoxQuestion.Width = (int)(img.Width * 0.2);
                    this.groupBoxQuestion.Height = groupBoxQuestion_Height + ((int)(img.Height * zoom) - this.pictureBoxQuestion.Height);
                    if (this.panel3.Height<panel_Height+((int)(img.Height * zoom) - this.pictureBoxQuestion.Height))
                        this.panel3.Height = this.panel_Height + ((int)(img.Height * zoom) - this.pictureBoxQuestion.Height);

                    this.pictureBoxQuestion.Height = (int)(img.Height * zoom);
                    this.pictureBoxQuestion.Image = img;

                    this.pictureBox3.Location = new Point(this.pictureBox3.Location.X, this.groupBoxQuestion.Location.Y + this.groupBoxQuestion.Height + 10);

                    
                }
                else
                {
                    this.pictureBoxQuestion.Image = null;
                    this.pictureBoxQuestion.Height = pictureBoxQuestion_Height;
                    this.groupBoxQuestion.Height = groupBoxQuestion_Height;
                }
                if (!pointExercise.Answer.ToString().Equals(""))
                {
                    if (pointExercise.Answer.ToString().Contains(".png"))
                    {
                        this.textBoxAnswer.Text = "";
                        Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + pointExercise.Answer);
                        zoom = (double)this.pictureBoxAnswer.Width / (double)img.Width;
                        marginTopAdd = ((int)(img.Height * zoom) - this.pictureBoxAnswer.Height);
                        this.groupBoxAnswer2.Height = groupBoxAnswer2_Height + ((int)(img.Height * zoom) - this.pictureBoxAnswer.Height);
                        
                        if (this.panel3.Height < panel_Height + ((int)(img.Height * zoom) - this.pictureBoxAnswer.Height))
                            this.panel3.Height = panel_Height + ((int)(img.Height * zoom) - this.pictureBoxAnswer.Height);

                        this.pictureBoxAnswer.Height = (int)(img.Height * zoom);
                        this.pictureBoxAnswer.Image = img;

                        this.pictureBox5.Location = new Point(this.pictureBox5.Location.X, this.groupBoxAnswer2.Location.Y + this.groupBoxAnswer2.Height + 10);
                    }
                    else
                    {
                        this.pictureBoxAnswer.Image = null;
                        this.textBoxAnswer.Text = pointExercise.Answer.ToString();

                        SetAnswerButton();
                    }
                }
                else
                {
                    //this.textBoxAnswer.Text = "";
                }
                
                if (!pointExercise.AnswerTips.ToString().Equals(""))
                {
                    Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + pointExercise.AnswerTips);
                    zoom = (double)660 / (double)img.Width;
                    marginTopAdd = ((int)(img.Height * zoom) - this.pictureBoxTips.Height);
                    //this.pictureBoxQuestion.Width = (int)(img.Width * 0.2);
                    this.groupBoxTips.Height = groupBoxTips_Height + ((int)(img.Height * zoom) - this.pictureBoxTips.Height);

                    if (this.panel4.Height < panel_Height + ((int)(img.Height * zoom) - this.pictureBoxTips.Height))
                        this.panel4.Height = panel_Height + ((int)(img.Height * zoom) - this.pictureBoxTips.Height);
                    this.pictureBoxTips.Height = (int)(img.Height * zoom);
                    this.pictureBoxTips.Image = img;

                    this.pictureBox7.Location = new Point(this.pictureBox7.Location.X, this.groupBoxTips.Location.Y + this.groupBoxTips.Height + 10);
                }
                else
                {
                    
                }
                if (!pointExercise.Analysis.ToString().Equals(""))
                {
                    Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + pointExercise.Analysis);
                    zoom = (double)660 / (double)img.Width;
                    marginTopAdd += ((int)(img.Height * zoom) - this.pictureBoxAnalysis.Height);
                    this.groupBoxAnalysis.Height = groupBoxAnalysis_Height + ((int)(img.Height * zoom) - this.pictureBoxAnalysis.Height);
                    if (this.panel4.Height < panel_Height + ((int)(img.Height * zoom) - this.pictureBoxAnalysis.Height))
                        this.panel4.Height = panel_Height + ((int)(img.Height * zoom) - this.pictureBoxAnalysis.Height);


                    this.pictureBoxAnalysis.Height = (int)(img.Height * zoom);
                    this.pictureBoxAnalysis.Image = img;
                    
                    this.pictureBox6.Location = new Point(this.pictureBox6.Location.X, this.groupBoxAnalysis.Location.Y + this.groupBoxAnalysis.Height + 10);
                }
                

                LoadKonwlegeTree("");
                this.textBox6.ForeColor = Color.LightGray;
                this.textBox6.Text = "输入关键字";
                
                if(pointExercise.KnowlegeList.Count!=0)
                {

                    for(int i=0;i<pointExercise.KnowlegeList.Count;i++)
                    {
                        ListViewItem item = new ListViewItem();
                        item.Name = pointExercise.KnowlegeList[i].Id.ToString();
                        item.Text = pointExercise.KnowlegeList[i].Name.ToString();
                        this.listView1.Items.Add(item);

                        
                        AddSelectedNodeTab(this.treeView1, pointExercise.KnowlegeList[i].Id.ToString());
                    }
                    
                }

            }
        }

        private Exercise GetExerciseByOrder(List<PaperNode> paperNodeList)
        {
            for(int i=0;i<paperNodeList.Count;i++)
            {
                if(exerciseOrder==exerciseOrderCount)
                {
                    Exercise reExercise = new Exercise();
                    reExercise.Question = (paperNodeList[i] as PaperNode).Title;
                    return reExercise;
                }
                else
                {
                    exerciseOrderCount += 1;
                    if (paperNodeList[i].ExerciseList.Count != 0)
                    {
                        Exercise exerciseTemp= GetExerciseByOrder(paperNodeList[i].ExerciseList as List<Exercise>);
                        if(exerciseTemp.Question!=null&&!exerciseTemp.Question.Equals(""))
                            return exerciseTemp;
                    }
                    else if (paperNodeList[i].PaperNodeList.Count != 0)
                    {
                        Exercise exerciseTemp = GetExerciseByOrder(paperNodeList[i].PaperNodeList as List<PaperNode>);
                        if (exerciseTemp.Question != null && !exerciseTemp.Question.Equals(""))
                            return exerciseTemp;
                    }
                }
            }
            return new Exercise();
        }

        private Exercise GetExerciseByOrder(List<Exercise> exerciseList)
        {
            for(int i=0;i<exerciseList.Count;i++)
            {
                if(exerciseOrder==exerciseOrderCount)
                {
                    return exerciseList[i] as Exercise;
                }
                else
                {
                    exerciseOrderCount += 1;
                    if (exerciseList[i].ExerciseList.Count != 0)
                    {
                        Exercise exerciseTemp=GetExerciseByOrder(exerciseList[i].ExerciseList as List<Exercise>);
                        if (exerciseTemp.Question != null && !exerciseTemp.Question.Equals(""))
                            return exerciseTemp;
                    }
                }
            }
            return new Exercise();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            foreach (Form f in Application.OpenForms)
            {
                if (f is PapersList)
                {
                    f.WindowState = FormWindowState.Maximized;
                }
            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            bool result = false;
            if (this.comboBoxType.SelectedIndex > -1)
            {
                pointExercise.Type = this.comboBoxType.SelectedIndex + 1;
            }
            if (this.comboBox1.SelectedIndex > -1)
            {
                pointExercise.AnswerNumber = this.comboBox1.SelectedIndex + 1;
            }
            if(pointExercise.Type==1||pointExercise.Type==2||pointExercise.Type==3)
            {
                string resultAnswer = "";
                foreach (Control c in this.groupBoxAnswer.Controls)
                {
                    if (c is Button)
                    {
                        Button button = c as Button;
                        if (button.BackColor == Color.LightGray)
                        {
                            resultAnswer = resultAnswer + button.Name.ToString() + ",";
                        }
                    }
                }
                if (!resultAnswer.Equals(""))
                    pointExercise.Answer = resultAnswer.Substring(0, resultAnswer.Length - 1);
            }
            

            int output;
            if (int.TryParse(this.textBox4.Text.ToString(), out output))
            {
                pointExercise.Score = int.Parse(this.textBox4.Text.ToString());
            }
            else
            {
                MessageBox.Show("分数格式不正确");
            }
            if (int.TryParse(this.textBox3.Text.ToString(), out output))
            {
                pointExercise.PredictDifficult = int.Parse(this.textBox3.Text.ToString());
            }
            else
            {
                MessageBox.Show("难度格式不正确");
            }
            if (int.TryParse(this.textBox5.Text.ToString(), out output))
            {
                pointExercise.Split = int.Parse(this.textBox5.Text.ToString());
            }
            else
            {
                MessageBox.Show("区分度格式不正确");
            }

            pointExercise.KnowlegeList.Clear();
            foreach(ListViewItem item in this.listView1.Items)
            {
                Knowlege temp = new Knowlege();
                temp.Id = int.Parse(item.Name);
                temp.Name = item.Text;
                pointExercise.KnowlegeList.Add(temp);
                AddSelectedNodeTab(this.treeView1,temp.Id.ToString());
            }

            if(this.textBox1.ForeColor!=Color.LightGray)
                pointExercise.Source = this.textBox1.Text.ToString();


            if (paper.PaperNodeList.Count != 0 && problemSet.ExerciseList.Count == 0)
            {
                exerciseOrderCount = 0;
                result = UpdateExercise(paper.PaperNodeList);
            }
            else if (paper.PaperNodeList.Count == 0 && problemSet.ExerciseList.Count != 0)
            {
                exerciseOrderCount = 0;
                result = UpdateExercise(problemSet.ExerciseList);
            }
            if(result)
            {
                if (paper.PaperNodeList.Count != 0 && problemSet.ExerciseList.Count == 0)
                {
                    jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(paper), Globals.ThisAddIn.exerciseJsonPath + paperName + ".json");
                }
                else if (paper.PaperNodeList.Count == 0 && problemSet.ExerciseList.Count != 0)
                {
                    jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(problemSet), Globals.ThisAddIn.exerciseJsonPath + paperName + ".json");
                }

                MessageBox.Show("保存成功");

                //显示下一题
                exerciseOrder += 1;

                Exercise tempEx = new Exercise();
                exerciseOrderCount = 0;
                if (paper.PaperNodeList.Count != 0 && problemSet.ExerciseList.Count == 0)
                {
                    tempEx = GetExerciseByOrder(paper.PaperNodeList as List<PaperNode>);
                }
                else if (paper.PaperNodeList.Count == 0 && problemSet.ExerciseList.Count != 0)
                {
                    tempEx = GetExerciseByOrder(problemSet.ExerciseList as List<Exercise>);
                }
                if (tempEx.Question.Contains(".png"))
                {
                    pointExercise = tempEx;
                    ReSetQuestionDetail();
                    this.label11.BackColor = Color.White;
                    this.label11.ForeColor = Color.Blue;
                    this.label10.BackColor = Control.DefaultBackColor;
                    this.label10.ForeColor = Label.DefaultForeColor;
                    this.label12.BackColor = Control.DefaultBackColor;
                    this.label12.ForeColor = Label.DefaultForeColor;
                    this.label13.BackColor = Control.DefaultBackColor;
                    this.label13.ForeColor = Label.DefaultForeColor;
                    SetTabVisible();

                    this.ScrollControlIntoView(this.groupBoxQuestion);

                }
                else
                {
                    MessageBox.Show("无更多题目");
                    exerciseOrder -= 1;

                    foreach (Form f in Application.OpenForms)
                    {
                        if (f is PapersList)
                        {
                            f.WindowState = FormWindowState.Maximized;
                        }
                    }
                }
                
                
            }
        }

        private bool UpdateExercise(List<PaperNode> paperNodeList)
        {
            for(int i=0; i<paperNodeList.Count;i++)
            {
                if(exerciseOrder==exerciseOrderCount)
                {
                    paperNodeList[i].Title = pointExercise.Question;
                    return true;
                }
                else
                {
                    exerciseOrderCount+=1;
                    if(paperNodeList[i].PaperNodeList.Count!=0)
                    {
                        bool result = UpdateExercise(paperNodeList[i].PaperNodeList as List<PaperNode>);
                        if (result)
                            return true;
                    }
                    if(paperNodeList[i].ExerciseList.Count!=0)
                    {
                        bool result = UpdateExercise(paperNodeList[i].ExerciseList as List<Exercise>);
                        if (result)
                            return true;
                    }
                }
            }
            return false;
        }

        private bool UpdateExercise(List<Exercise> exerciseList)
        {
            for (int i = 0; i < exerciseList.Count; i++)
            {
                if (exerciseOrder == exerciseOrderCount)
                {
                    exerciseList[i] = pointExercise;
                    return true;
                }
                else
                {
                    exerciseOrderCount += 1;
                    if(exerciseList[i].ExerciseList.Count!=0)
                    {
                        bool result = UpdateExercise(exerciseList[i].ExerciseList as List<Exercise>);
                        if (result)
                            return true;
                    }
                }
            }
            return false;
        }

        private void buttonQuestion_Click(object sender, EventArgs e)
        {
            this.panel1.Focus();
            if (Globals.ThisAddIn.Application.Selection.Range.Text != null)
            {
                string imgName = imgHelper.GetSelectionImg(paperName);
                pointExercise.Question = imgName;


                this.pictureBoxQuestion.Image = null;
                this.pictureBoxQuestion.Height = pictureBoxQuestion_Height;
                this.groupBoxQuestion.Height = groupBoxQuestion_Height;

                Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + pointExercise.Question);
                double zoom = (double)this.pictureBoxQuestion.Width / (double)img.Width;
                int marginTopAdd = ((int)(img.Height * zoom) - this.pictureBoxQuestion.Height);
                //this.pictureBoxQuestion.Width = (int)(img.Width * 0.2);
                this.groupBoxQuestion.Height = groupBoxQuestion_Height + ((int)(img.Height * zoom) - this.pictureBoxQuestion.Height);
                this.panel3.Height = this.panel_Height + ((int)(img.Height * zoom) - this.pictureBoxQuestion.Height);

                this.pictureBoxQuestion.Height = (int)(img.Height * zoom);
                this.pictureBoxQuestion.Image = img;

                this.pictureBox3.Location = new Point(this.pictureBox3.Location.X, this.groupBoxQuestion.Location.Y + this.groupBoxQuestion.Height + 10);

                //ReSetQuestionDetail();
            }
            
        }

        private void buttonAnswer1_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Range.Text != null)
            {
                string selectionText = Globals.ThisAddIn.Application.Selection.Range.Text.ToString();
                pointExercise.Answer = selectionText;

                ReSetQuestionDetail();
            }
        }

        private void buttonAnswer2_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Range.Text != null)
            {
                this.textBoxAnswer.Text = "";


                string imgName = imgHelper.GetSelectionImg(paperName);
                pointExercise.Answer = imgName;
                pointExercise.Type = this.comboBoxType.SelectedIndex + 1;

                this.pictureBoxAnswer.Height = pictureBoxAnswer_Height;
                this.pictureBoxAnswer.Image = null;
                this.groupBoxAnswer2.Height = groupBoxAnswer2_Height;

                Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + pointExercise.Answer);
                double  zoom = (double)this.pictureBoxAnswer.Width / (double)img.Width;
                int marginTopAdd = ((int)(img.Height * zoom) - this.pictureBoxAnswer.Height);
                this.groupBoxAnswer2.Height = groupBoxAnswer2_Height + ((int)(img.Height * zoom) - this.pictureBoxAnswer.Height);

                if (this.panel3.Height < panel_Height + ((int)(img.Height * zoom) - this.pictureBoxAnswer.Height))
                    this.panel3.Height = panel_Height + ((int)(img.Height * zoom) - this.pictureBoxAnswer.Height);

                this.pictureBoxAnswer.Height = (int)(img.Height * zoom);
                this.pictureBoxAnswer.Image = img;

                this.pictureBox5.Location = new Point(this.pictureBox5.Location.X, this.groupBoxAnswer2.Location.Y + this.groupBoxAnswer2.Height + 10);

                //ReSetQuestionDetail();
            }
        }

        private void buttonAnalysis_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Range.Text != null)
            {

                string imgName = imgHelper.GetSelectionImg(paperName);

                pointExercise.Analysis = imgName;

                this.pictureBoxAnalysis.Height = pictureBoxAnalysis_Height;
                this.pictureBoxAnalysis.Image = null;
                this.groupBoxAnalysis.Height = groupBoxAnalysis_Height;

                Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + pointExercise.Analysis);
                double zoom = (double)660 / (double)img.Width;
                int marginTopAdd = ((int)(img.Height * zoom) - this.pictureBoxAnalysis.Height);
                this.groupBoxAnalysis.Height = groupBoxAnalysis_Height + ((int)(img.Height * zoom) - this.pictureBoxAnalysis.Height);
                if (this.panel4.Height < panel_Height + ((int)(img.Height * zoom) - this.pictureBoxAnalysis.Height))
                    this.panel4.Height = panel_Height + ((int)(img.Height * zoom) - this.pictureBoxAnalysis.Height);


                this.pictureBoxAnalysis.Height = (int)(img.Height * zoom);
                this.pictureBoxAnalysis.Image = img;

                this.pictureBox6.Location = new Point(this.pictureBox6.Location.X, this.groupBoxAnalysis.Location.Y + this.groupBoxAnalysis.Height + 10);


                //ReSetQuestionDetail();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            openFileDialog1.Filter = "mp4文件(*.mp4)|*.mp4|所有文件(*)|*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string PicFilePath = openFileDialog1.FileName;
                this.textBox2.Text = PicFilePath;

                string[] videoSplit = PicFilePath.Split('\\');
                string videoName=videoSplit[videoSplit.Length-1];

                pointExercise.Video = videoName;

                File.Copy(PicFilePath,Globals.ThisAddIn.exerciseJsonPath+paperName+"\\"+videoName,true);
            }
            else if(openFileDialog1.ShowDialog()==DialogResult.Cancel)
            {
                this.textBox2.Text = "";
                pointExercise.Video = "";

            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Range.Text != null)
            {
                string imgName = imgHelper.GetSelectionImg(paperName);

                pointExercise.AnswerTips = imgName;

                this.pictureBoxTips.Height = pictureBoxTips_Height;
                this.pictureBoxTips.Image = null;
                this.groupBoxTips.Height = groupBoxTips_Height;

                Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + pointExercise.AnswerTips);
                double zoom = (double)660 / (double)img.Width;
                int marginTopAdd = ((int)(img.Height * zoom) - this.pictureBoxTips.Height);
                //this.pictureBoxQuestion.Width = (int)(img.Width * 0.2);
                this.groupBoxTips.Height = groupBoxTips_Height + ((int)(img.Height * zoom) - this.pictureBoxTips.Height);

                if (this.panel4.Height < panel_Height + ((int)(img.Height * zoom) - this.pictureBoxTips.Height))
                    this.panel4.Height = panel_Height + ((int)(img.Height * zoom) - this.pictureBoxTips.Height);
                this.pictureBoxTips.Height = (int)(img.Height * zoom);
                this.pictureBoxTips.Image = img;

                this.pictureBox7.Location = new Point(this.pictureBox7.Location.X, this.groupBoxTips.Location.Y + this.groupBoxTips.Height + 10);


                //ReSetQuestionDetail();
            }
        }

        private void label11_Click(object sender, EventArgs e)
        {
            Label label = sender as Label;
            label.BackColor = Color.White;
            label.ForeColor = Color.Blue;

            foreach (Control c in this.panel2.Controls)
            {
                if (c is Label && c.Name != label.Name)
                {
                    c.BackColor = Control.DefaultBackColor;
                    c.ForeColor = Label.DefaultForeColor;
                }
            }
            SetTabVisible();
        }

        private void SetTabVisible()
        {
            this.panel3.Visible=(this.label11.BackColor==Color.White);
            this.panel4.Visible = (this.label10.BackColor == Color.White);
            this.panel5.Visible = (this.label12.BackColor == Color.White);
            this.panel6.Visible = (this.label13.BackColor == Color.White);
        }

        private void panel3_Click(object sender, EventArgs e)
        {
            this.panel1.Focus();
        }

        private void pictureBoxNext_Click(object sender, EventArgs e)
        {
            //显示下一题
            exerciseOrder += 1;

            Exercise tempEx = new Exercise();
            exerciseOrderCount = 0;
            if (paper.PaperNodeList.Count != 0 && problemSet.ExerciseList.Count == 0)
            {
                tempEx = GetExerciseByOrder(paper.PaperNodeList as List<PaperNode>);
            }
            else if (paper.PaperNodeList.Count == 0 && problemSet.ExerciseList.Count != 0)
            {
                tempEx = GetExerciseByOrder(problemSet.ExerciseList as List<Exercise>);
            }
            if (tempEx.Question.Contains(".png"))
            {
                pointExercise = tempEx;
                ReSetQuestionDetail();

                this.label11.BackColor = Color.White;
                this.label11.ForeColor = Color.Blue;
                this.label10.BackColor = Control.DefaultBackColor;
                this.label10.ForeColor = Label.DefaultForeColor;
                this.label12.BackColor = Control.DefaultBackColor;
                this.label12.ForeColor = Label.DefaultForeColor;
                this.label13.BackColor = Control.DefaultBackColor;
                this.label13.ForeColor = Label.DefaultForeColor;
                SetTabVisible();
                this.ScrollControlIntoView(this.groupBoxQuestion);

            }
            else
            {
                MessageBox.Show("无更多题目");
                exerciseOrder -= 1;
            }
        }

        private void pictureBoxPre_Click(object sender, EventArgs e)
        {
            //显示上一题
            if(exerciseOrder>0)
            {
                exerciseOrder -= 1;

                Exercise tempEx = new Exercise();
                exerciseOrderCount = 0;
                if (paper.PaperNodeList.Count != 0 && problemSet.ExerciseList.Count == 0)
                {
                    tempEx = GetExerciseByOrder(paper.PaperNodeList as List<PaperNode>);
                }
                else if (paper.PaperNodeList.Count == 0 && problemSet.ExerciseList.Count != 0)
                {
                    tempEx = GetExerciseByOrder(problemSet.ExerciseList as List<Exercise>);
                }
                if (tempEx.Question.Contains(".png"))
                {
                    pointExercise = tempEx;
                    ReSetQuestionDetail();
                    this.label11.BackColor = Color.White;
                    this.label11.ForeColor = Color.Blue;
                    this.label10.BackColor = Control.DefaultBackColor;
                    this.label10.ForeColor = Label.DefaultForeColor;
                    this.label12.BackColor = Control.DefaultBackColor;
                    this.label12.ForeColor = Label.DefaultForeColor;
                    this.label13.BackColor = Control.DefaultBackColor;
                    this.label13.ForeColor = Label.DefaultForeColor;
                    SetTabVisible();
                    this.ScrollControlIntoView(this.groupBoxQuestion);

                }
                else
                {
                    MessageBox.Show("无更多题目");
                    exerciseOrder -= 1;
                }
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode node = this.treeView1.SelectedNode;
            node.BackColor = Color.DarkGray;
            bool ISEXIT = false;
            foreach(ListViewItem item in this.listView1.Items)
            {
                if(item.Name==node.Name)
                {
                    ISEXIT = true;
                }
            }
            if(!ISEXIT)
            {
                ListViewItem item = new ListViewItem();
                item.Name = node.Name;
                item.Text = node.Text;
                this.listView1.Items.Add(item);
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem item in this.listView1.SelectedItems)
            {
                this.listView1.Items.Remove(item);
                RemoveSelectedNodeTab(this.treeView1,item.Name);
            }
        }

        private void AddSelectedNodeTab(TreeView tree, string Name)
        {
            foreach (TreeNode node in tree.Nodes)
            {
                if (node.Name.Equals(Name))
                {
                    node.BackColor = Color.DarkGray;
                }
                if (node.Nodes.Count != 0)
                {
                    AddSelectedNodeTab(node.TreeView, Name);
                }
            }
        }

        private void RemoveSelectedNodeTab(TreeView tree, string Name)
        {
            foreach(TreeNode node in tree.Nodes)
            {
                if(node.Name.Equals(Name))
                {
                    node.BackColor = Color.White;
                }
                if(node.Nodes.Count!=0)
                {
                    RemoveSelectedNodeTab(node.TreeView,Name);
                }
            }
        }

        private void RemoveAllSelectedNodeTab(TreeView tree)
        {
            foreach (TreeNode node in tree.Nodes)
            {
                node.BackColor = Color.White;
                if (node.Nodes.Count != 0)
                {
                    RemoveAllSelectedNodeTab(node.TreeView);
                }
            }
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            if(this.textBox1.ForeColor==Color.LightGray)
            {
                this.textBox1.Text = "";
                this.textBox1.ForeColor = Color.Black;
            }
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            if(this.textBox6.ForeColor!=Color.LightGray)
            {
                LoadKonwlegeTree(this.textBox6.Text.ToString());
                for (int i = 0; i < this.listView1.Items.Count; i++)
                {
                    ListViewItem item = this.listView1.Items[i] as ListViewItem;
                    AddSelectedNodeTab(this.treeView1, item.Name.ToString());
                }
            }
            else
            {

            }
        }

        private void textBox6_Click(object sender, EventArgs e)
        {
            if(this.textBox6.ForeColor==Color.LightGray)
            {
                this.textBox6.Text = "";
                this.textBox6.ForeColor = Color.Black;
            }
        }

        private void SetAnswerButton()
        {
            this.groupBoxAnswer.Controls.Clear();

            if (this.comboBoxType.SelectedIndex + 1 == 1 || this.comboBoxType.SelectedIndex + 1 == 2)
            {
                this.label5.Visible = true;
                this.comboBox1.Visible = true;
                this.groupBoxAnswer.Visible = true;
                this.groupBoxAnswer2.Visible = false;
                this.pictureBox5.Visible = false;

                char answer = 'A';
                for (int i = 0; i < this.comboBox1.SelectedIndex + 1; i++)
                {
                    Button button = new Button();
                    button.BackColor = Color.White;
                    button.Name = answer + "";
                    button.Text = answer + "";
                    button.Size = new Size(42, 42);
                    button.Location = new Point((i + 1) * 50, 27);
                    button.Click += new EventHandler(LetterButton_Click);
                    this.groupBoxAnswer.Controls.Add(button);
                    answer += '\u0001';
                }
            }
            else if (this.comboBoxType.SelectedIndex + 1 == 3)
            {
                this.label5.Visible = false;
                this.comboBox1.Visible = false;
                this.groupBoxAnswer.Visible = true;
                this.groupBoxAnswer2.Visible = false;
                this.pictureBox5.Visible = false;

                Button buttonF = new Button();
                buttonF.BackColor = Color.White;
                buttonF.Name = "0";
                buttonF.Text = "错";
                buttonF.Location = new Point(50, 27);
                buttonF.Size = new Size(42, 42);
                buttonF.Click += new EventHandler(PanDuanButton_Click);
                this.groupBoxAnswer.Controls.Add(buttonF);

                Button buttonT = new Button();
                buttonT.BackColor = Color.White;
                buttonT.Name = "1";
                buttonT.Text = "对";
                buttonT.Location = new Point(100, 27);
                buttonT.Size = new Size(42, 42);
                buttonT.Click += new EventHandler(PanDuanButton_Click);
                this.groupBoxAnswer.Controls.Add(buttonT);
            }
            else if (this.comboBoxType.SelectedIndex + 1 == 4 || this.comboBoxType.SelectedIndex + 1 == 5)
            {
                this.label5.Visible = false;
                this.comboBox1.Visible = false;
                this.groupBoxAnswer.Visible = false;
                this.groupBoxAnswer2.Visible = true;
                this.pictureBox5.Visible = true;
            }

            foreach(Control c in this.groupBoxAnswer.Controls)
            {
                if (pointExercise.Answer.ToString().Contains(c.Name.ToString()))
                    c.BackColor = Color.LightGray;
            }
        }

        private void comboBoxType_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetAnswerButton();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetAnswerButton();
        }

        private void LetterButton_Click(object sender, EventArgs e)
        {
            if(sender is Button)
            {
                Button b = sender as Button;
                if (this.comboBoxType.SelectedIndex + 1 == 1)
                {
                    b.BackColor = Color.LightGray;
                    foreach (Control c in this.groupBoxAnswer.Controls)
                    {
                        if (c is Button)
                        {
                            Button button = c as Button;
                            if (!button.Name.ToString().Equals(b.Name.ToString()))
                            {
                                button.BackColor = Color.White;
                            }
                        }
                    }
                }
                else if (this.comboBoxType.SelectedIndex + 1 == 2)
                {
                    if (b.BackColor == Color.LightGray)
                        b.BackColor = Color.White;
                    else if (b.BackColor == Color.White)
                        b.BackColor = Color.LightGray;
                }
                
            }
        }
        

        private void PanDuanButton_Click(object sender, EventArgs e)
        {
            if(sender is Button)
            {
                Button b = sender as Button;
                foreach(Control c in this.groupBoxAnswer.Controls)
                {
                    if(c is Button)
                    {
                        Button button = c as Button;
                        if(button.Text.ToString().Equals(b.Text.ToString()))
                        {
                            button.BackColor = Color.LightGray;
                        }
                        else
                        {
                            button.BackColor = Color.White;
                        }
                    }
                }
            }
        }
    }
}
