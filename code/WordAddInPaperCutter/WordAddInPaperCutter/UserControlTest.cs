using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Drawing.Drawing2D;
using WordAddInPaperCutter.Common;
using System.Windows.Threading;

namespace WordAddInPaperCutter
{
    public partial class UserControlTest : UserControl
    {


        private ImgHelper imgHelper = new ImgHelper();
        private DispatcherTimer timer = new DispatcherTimer();
        private DispatcherTimer timerForProgress = new DispatcherTimer();

        private int pictureBox1_Height;
        private int groupBox1_Height;
        string paperNameReal = "";
        //string paperName = "";
        string path = "";

        public UserControlTest()
        {
            InitializeComponent();

            pictureBox1_Height = this.pictureBox1.Height;
            groupBox1_Height = this.groupBox1.Height;

            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += theout;

            timerForProgress.Interval = TimeSpan.FromSeconds(0.1);
            timerForProgress.Tick += theoutForProgress;

            this.label5.Visible = true;
            this.label5.Text = "";
        }

        private void theout(object source, EventArgs e)
        {

            this.label3.Visible = false;
            timer.Stop();
        }

        private void theoutForProgress(object source, EventArgs e)
        {
            if(imgHelper.cutTimes>5)
            {
                this.progressBar1.Visible = true;
                this.progressBar1.Maximum = imgHelper.cutTimes;
                this.progressBar1.Value = imgHelper.cutTimesCount;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Label label = sender as Label;
            label.BackColor = Color.White;
            label.ForeColor = Color.Blue;
           
            foreach(Control c in this.panel1.Controls)
            {
                if(c is Label&&c.Name!=label.Name)
                {
                    c.BackColor = Control.DefaultBackColor;
                    c.ForeColor = Label.DefaultForeColor;
                }
            }
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            SetSavePath();
        }

        private void SetSavePath()
        {
            
            System.Windows.Forms.FolderBrowserDialog openFileDialog1 = new System.Windows.Forms.FolderBrowserDialog();

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string PicFilePath = openFileDialog1.SelectedPath.ToString();
                this.textBox2.Text = PicFilePath;
            }
            else if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                this.textBox2.Text = "";
            }
        }

        private void SetProgressBarProgress()
        {
            //imgHelper.SetProgressBarProgress(this.progressBar1);
            while(imgHelper.inProgress)
            {
                this.progressBar1.Invoke((EventHandler)delegate
                {
                    this.progressBar1.Maximum = imgHelper.cutTimes;
                    this.progressBar1.Value = imgHelper.cutTimesCount;
                });
                this.label5.Invoke((EventHandler)delegate
                {
                    this.label5.Text = imgHelper.progress + "...";
                });
                System.Threading.Thread.Sleep(1);
            }
            this.label5.Invoke((EventHandler)delegate
            {
                this.label5.Text = "";
            });
        }

        private void GetImgName()
        {
            imgHelper.GetSelectionImg(path,paperNameReal);

            imgHelper.inProgress = false;

            this.progressBar1.Invoke((EventHandler)delegate
            {
                this.progressBar1.Visible = false;
                this.progressBar1.Value = 0;
            });

            string[] allImg = Directory.GetFiles(Globals.ThisAddIn.exerciseJsonPath + "Temp3\\" + paperNameReal);
            if (allImg.Length == 0)
                return;
            Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + "Temp3\\" + paperNameReal + "\\0.png");
            double beishu = (double)660 / (double)img.Width;
            this.groupBox1.Invoke((EventHandler)delegate
            {
                int heightAdd = 0;
                for (int i =0; i < allImg.Length;i++ )
                {
                    Image otherImg = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + "Temp3\\" + paperNameReal+"\\"+i+".png");

                    heightAdd += (int)(otherImg.Height*beishu);
                }
                this.groupBox1.Height = heightAdd+10;

                heightAdd = 0;
                for (int i = 0; i < allImg.Length; i++)
                {
                    Image otherImg = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + "Temp3\\" + paperNameReal + "\\" + i + ".png");

                    PictureBox pb = new PictureBox();
                    pb.Image = otherImg;
                    pb.SizeMode = PictureBoxSizeMode.Zoom;
                    pb.Anchor = AnchorStyles.Right | AnchorStyles.Left | AnchorStyles.Top;
                    pb.Size = new Size(660, (int)(otherImg.Height * beishu));
                    pb.Location = new System.Drawing.Point(3, heightAdd);
                    this.groupBox1.Controls.Add(pb);

                    heightAdd += pb.Height;
                }

                this.groupBox1.Height = heightAdd+10;
            });
        }


        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Range.Text != null)
            {
                paperNameReal=Guid.NewGuid().ToString();
                this.groupBox1.Controls.Clear();

                this.label1.Visible = false;
                imgHelper.inProgress = true;
                this.progressBar1.Visible = true;
                System.Threading.Thread t1 = new System.Threading.Thread(GetImgName);
                t1.Start();
                System.Threading.Thread.Sleep(3);
                System.Threading.Thread t = new System.Threading.Thread(SetProgressBarProgress);
                t.Start();
                ////timerForProgress.Start();
                //imgName = imgHelper.GetSelectionImg("exerciseUnClassified");
                ////timerForProgress.Stop();
                ////this.progressBar1.Value = 0;
                ////this.progressBar1.Visible = false;

                //t.Abort();

                //this.progressBar1.Visible = true;
                //while (t1.IsAlive)
                //{
                //    this.progressBar1.Invoke((EventHandler)delegate
                //    {
                //        this.progressBar1.Maximum = imgHelper.cutTimes;
                //        this.progressBar1.Value = imgHelper.cutTimesCount;
                //        //System.Threading.Thread.Sleep(1);
                        
                //    });
                //}
                //this.progressBar1.Visible = false;

                //while(t1.IsAlive)
                //{
                //    System.Threading.Thread.Sleep(1);
                //}
                //t.Abort();

                    
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.SaveFileDialog openFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            openFileDialog1.Filter = "zip文件(*.zip)|*.zip";

            string savePath = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                savePath = openFileDialog1.FileName;
            }
            else if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            ZipClass zip = new ZipClass();
            zip.ZipFileFromDirectory(Globals.ThisAddIn.exerciseJsonPath + "Temp3\\" + paperNameReal, savePath, 0);
            this.label3.Visible = true;
            timer.Start();
        }


    }
}
