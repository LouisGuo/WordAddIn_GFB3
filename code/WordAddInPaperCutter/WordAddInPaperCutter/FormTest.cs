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
using Newtonsoft.Json;
using WordAddInPaperCutter.JsonClass;

namespace WordAddInPaperCutter
{
    public partial class FormTest : Form
    {
        private JsonFileHelper jsonFileHelper = new JsonFileHelper();
        public FormTest()
        {
            InitializeComponent();
        }

        private void FormTest_Load(object sender, EventArgs e)
        {
            string str = jsonFileHelper.GetFileString(@"D:\新建文件夹\imgJsonList.json");

            List<ExerciseWithRange> rangeExerciseList = new List<ExerciseWithRange>();
            rangeExerciseList = (List<ExerciseWithRange>)JsonConvert.DeserializeObject(str, rangeExerciseList.GetType());
            ProblemSet problemSet = new ProblemSet();
            problemSet = jsonFileHelper.GetProblemSetFromFile(Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified.json");

            List<Exercise> ExerciseList = problemSet.ExerciseList;

            int oldImgLocation_x = 0;
            int oldImgWidth = 660;
            int marginTop = 0;
            if(rangeExerciseList.Count!=0)
            {
                for(int i=0;i<rangeExerciseList.Count;i++)
                {
                    Image img = Image.FromFile(@"D:\新建文件夹\"+rangeExerciseList[i].Question);
                    marginTop += img.Height;
                    
                    
                }
                this.panel1.Height = marginTop + 100;

                marginTop = 0;
                for (int i = 0; i < rangeExerciseList.Count; i++)
                {
                    Image img = Image.FromFile(@"D:\新建文件夹\" + rangeExerciseList[i].Question);
                    PictureBox pictureBox = new PictureBox();
                    pictureBox.Image = img;
                    pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
                    pictureBox.Size = new System.Drawing.Size(img.Width,img.Height);
                    pictureBox.Location = new Point(0, marginTop);

                    this.panel1.Controls.Add(pictureBox);
                    marginTop += img.Height;

                    if (img.Width > oldImgWidth)
                        oldImgWidth = img.Width;

                    if (pictureBox.Location.X + pictureBox.Width > oldImgLocation_x)
                        oldImgLocation_x = pictureBox.Location.X + pictureBox.Width;

                }
            }

            oldImgLocation_x += 10;
            marginTop = 0;
            if(ExerciseList.Count!=0)
            {
                for (int i = 0; i < ExerciseList.Count; i++)
                {
                    Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified\\" + ExerciseList[i].Question);
                    
                    marginTop += (int)(((double)oldImgWidth / (double)img.Width) * img.Height);
                }

                if(this.panel1.Height<marginTop)
                    this.panel1.Height = marginTop + 100;
                marginTop = 0;

                for(int i=0;i<ExerciseList.Count;i++)
                {
                    Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified\\" + ExerciseList[i].Question);
                    PictureBox pictureBox = new PictureBox();
                    pictureBox.Image = img;
                    pictureBox.SizeMode = PictureBoxSizeMode.Zoom;


                    pictureBox.Size = new System.Drawing.Size(oldImgWidth, (int)(((double)oldImgWidth/(double)img.Width)*img.Height));
                    pictureBox.Location = new Point(oldImgLocation_x, marginTop);

                    this.panel1.Controls.Add(pictureBox);
                    marginTop += (int)(((double)oldImgWidth / (double)img.Width) * img.Height);
                }
            }
            
        }
    }
}
