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
using System.Drawing.Imaging;
using System.IO;
using Newtonsoft.Json;
using System.Windows.Threading;
using System.Drawing.Drawing2D;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace WordAddInPaperCutter
{
    public partial class UserControl1 : UserControl
    {
        private DB db = new DB();
        private List<ExerciseWithRange> rangeList = new List<ExerciseWithRange>();

        private DispatcherTimer timer = new DispatcherTimer();
        private JsonFileHelper jsonFileHelper = new JsonFileHelper();
        private ImgHelper imgHelper = new ImgHelper();

        private Exercise exerciseTemp = new Exercise();
        private ProblemSet exerciseUnClassified = new ProblemSet();

        private int pictureBox1_Height;
        private int groupBox1_Height;
        private int panle1_Height;


        public UserControl1()
        {
            InitializeComponent();

            try
            {
                //jsonFileHelper.WriteFileString("", @"D:\新建文件夹\imgJsonList.json");
                //rangeList = (List<Exercise>)JsonConvert.DeserializeObject(jsonFileHelper.GetFileString(@"D:\新建文件夹\imgJsonList.json"), rangeList.GetType());

                pictureBox1_Height = this.pictureBox1.Height;
                groupBox1_Height = this.groupBox1.Height;
                panle1_Height = this.panel1.Height;

                timer.Interval = TimeSpan.FromSeconds(3);
                timer.Tick += theout;

                this.comboBox1.SelectedIndex = 0;

                //paperID = System.Guid.NewGuid().ToString();

                //DateTime now = DateTime.Now;
                //paperID = now.Year+"-"+now.Month+"-"+now.Day+" "+now.Hour+"."+now.Minute+"."+now.Second+"__"+new Random().Next(1000,9999);

            }
            catch
            {
                MessageBox.Show("无法初始化工具一");
            }
            
        }

        
        private void GetSelectionPicture(int type)
        {
            if (Globals.ThisAddIn.Application.Selection.Range.Text != null)
            {
                //保存range数据


                this.label1.Visible = false;

                exerciseUnClassified = jsonFileHelper.GetProblemSetFromFile(Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified.json");

                string imgName=imgHelper.GetSelectionImg("exerciseUnClassified");
                if (imgName.Equals(""))
                    return;
                Image img = Image.FromFile(Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified\\"+imgName);


                
                
                this.textBox1.Visible = false;
                this.pictureBox1.Visible = true;
                //this.pictureBox1.Width = 660;
                double beishu = (double)660 / (double)img.Width;
                this.pictureBox1.Height = (int)(img.Height * beishu);


                this.panel1.Height = panle1_Height + this.pictureBox1.Height - pictureBox1_Height;

                this.pictureBox1.Image = img;

                exerciseTemp.Question = imgName;
                exerciseTemp.QuestionType = type;

                
                if (this.button7.Visible == true)
                {
                    ExerciseWithRange exRange = new ExerciseWithRange();
                    exRange.RangeStart = Globals.ThisAddIn.Application.Selection.Range.Start;
                    exRange.RangeEnd = Globals.ThisAddIn.Application.Selection.Range.End;
                    rangeList.Add(exRange);
                    try
                    {
                        jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(rangeList), @"D:\新建文件夹\imgJsonList.json");
                    }
                    catch
                    {

                    }
                }
                
                
                //exerciseUnClassified.ExerciseList.Add(exerciseTemp);

                //jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(exerciseUnClassified), Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified.json");
            }
        }

        
        private void theout(object source, EventArgs e)
        {
            this.button1.Enabled = true;
            this.button2.Enabled = true;
            this.button3.Enabled = true;
            this.button4.Enabled = true;

            this.label1.Visible = false;
            timer.Stop();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(Globals.ThisAddIn.Application.Selection.Range.Text!=null)
            {
                GetSelectionPicture(0);
                button1.BackColor = Color.Red;
                button2.BackColor = Control.DefaultBackColor;
                button3.BackColor = Control.DefaultBackColor;
                button4.BackColor = Control.DefaultBackColor;

                this.button1.Enabled = false;
                timer.Start();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Range.Text != null)
            {
                GetSelectionPicture(1);
                button1.BackColor = Control.DefaultBackColor;
                button2.BackColor = Color.Red;
                button3.BackColor = Control.DefaultBackColor;
                button4.BackColor = Control.DefaultBackColor;

                this.button2.Enabled = false;
                timer.Start();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Range.Text != null)
            {
                GetSelectionPicture(2);
                button1.BackColor = Control.DefaultBackColor;
                button2.BackColor = Control.DefaultBackColor;
                button3.BackColor = Color.Red;
                button4.BackColor = Control.DefaultBackColor;

                this.button3.Enabled = false;
                timer.Start();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Range.Text != null)
            {
                GetSelectionPicture(3);

                button1.BackColor = Control.DefaultBackColor;
                button2.BackColor = Control.DefaultBackColor;
                button3.BackColor = Control.DefaultBackColor;
                button4.BackColor = Color.Red;

                this.button4.Enabled = false;
                timer.Start();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            //new PapersList(editorName).Show();
            new PapersList().Show();

            //new ZipClass().UnZip(@"C:\Resources\Papers\2014-6-13 22.28.24__5969####1234.zip",@"C:\Users\Word\Desktop\2014-6-13 22.28.24__5969####1234");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            exerciseUnClassified.ExerciseList.Add(exerciseTemp);

            jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(exerciseUnClassified), Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified.json");

            button1.BackColor = Control.DefaultBackColor;
            button2.BackColor = Control.DefaultBackColor;
            button3.BackColor = Control.DefaultBackColor;
            button4.BackColor = Control.DefaultBackColor;

            this.label1.Visible = true;
            timer.Start();
        }

        private void label2_Click(object sender, EventArgs e)
        {
            this.pictureBox1.Height = pictureBox1_Height;
            this.panel1.Height = panle1_Height;
            this.pictureBox1.Image = null;

            Label label = sender as Label;
            label.BackColor = Color.White;
            label.ForeColor = Color.Blue;

            this.panel1.ScrollControlIntoView(this.pictureBox1);

            foreach (Control c in this.panel2.Controls)
            {
                if (c is Label && c.Name != label.Name)
                {
                    c.BackColor = Control.DefaultBackColor;
                    c.ForeColor = Label.DefaultForeColor;
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Range.Text != null)
            {
                if(this.label2.BackColor==Color.White)
                {
                    GetSelectionPicture(0);
                }
                else if(this.label3.BackColor==Color.White)
                {
                    GetSelectionPicture(1);
                }

                this.panel1.Focus();
                timer.Start();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            exerciseUnClassified.ExerciseList.Add(exerciseTemp);
            jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(exerciseUnClassified), Globals.ThisAddIn.exerciseJsonPath + "exerciseUnClassified.json");

            this.label1.Visible = true;
            this.panel1.Focus();
            timer.Start();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            new PapersList().Show();
        }

        private void UserControl1_Scroll(object sender, ScrollEventArgs e)
        {
            
        }

        private void panel1_Click(object sender, EventArgs e)
        {
            this.panel3.Focus();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(this.comboBox1.SelectedIndex==0)
            {
                this.label2.Visible = true;
                label2.BackColor = Color.White;
                label2.ForeColor = Color.Blue;

                label3.BackColor = Control.DefaultBackColor;
                label3.ForeColor = Label.DefaultForeColor;


            }
            else if(this.comboBox1.SelectedIndex==1)
            {
                this.label2.Visible = false;

                label3.BackColor = Color.White;
                label3.ForeColor = Color.Blue;

                label2.BackColor = Control.DefaultBackColor;
                label2.ForeColor = Label.DefaultForeColor;
            }
        }
        

        
        private void GetBetterImg()
        {
            Range rangeAll = Globals.ThisAddIn.Application.ActiveDocument.Range();

            //range.SetRange(0,0);

            //while(range.End<rangeAll.End)
            //{
            //    range.Start = range.Start + 100;
            //    range.End = range.End + 100;
            //    Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(range);
            //}

            

            object filename = @"D:\BaoProject\高分宝三期\2014届高三数学理科模拟试题（定稿）.doc";  //文件保存路径
            Object Nothing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document WordDoc = WordApp.Documents.Add(ref filename, ref Nothing, ref Nothing, ref Nothing);


            Range range = WordDoc.Range(0, 500);



            range.Copy();



            rangeAll.Paste();

            rangeAll = Globals.ThisAddIn.Application.ActiveDocument.Range();

            rangeAll.SetRange(rangeAll.End, rangeAll.End);
            rangeAll.Select();

            System.Windows.Point pointend = CaretPos();

            rangeAll.SetRange(0, 0);
            rangeAll.Select();

            Microsoft.Office.Tools.Word.Document document = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);
            Microsoft.Office.Tools.Word.DropDownListContentControl dropdown = document.Controls.AddDropDownListContentControl(rangeAll, "MyContentControl");
            dropdown.PlaceholderText = "My DropdownList Test";
            dropdown.DropDownListEntries.Add("Test01", "01", 1);
            dropdown.DropDownListEntries.Add("Test02", "02", 2);
            dropdown.DropDownListEntries.Add("Test03", "03", 3);

            rangeAll.SetRange(dropdown.Range.End+1, dropdown.Range.End+1);

            System.Drawing.Point o = new System.Drawing.Point(dropdown.Application.Left, dropdown.Application.Top);

            System.Drawing.Point o1 = this.PointToScreen(o);


            rangeAll.SetRange(0, 0);
            rangeAll.Select();

            System.Drawing.Point currentPos = GetPositionForShowing(Globals.ThisAddIn.Application.Selection);
            //Globals.ThisAddIn._FloatingPanel = new FloatingPanel(bookmark);



            rangeAll.Text = rangeAll.Text + "\n" + dropdown.Application.Top + "::::" + dropdown.Application.Left + "\n" + o1.X + "::::::" + o1.Y + "\n" + currentPos.X + "::::::" + currentPos.Y;


            Bitmap image = new Bitmap(currentPos.X, currentPos.Y);

            Graphics g = Graphics.FromImage(image);

            System.Drawing.Point FrmP = new System.Drawing.Point(currentPos.X, currentPos.Y);
            //ScreenP返回相对屏幕的坐标
            System.Drawing.Point ScreenP = this.PointToScreen(FrmP);

            g.CopyFromScreen(0, 0, 0, 0, image.Size);

            image.Save(@"C:\Users\Word\Desktop\3523433.png", System.Drawing.Imaging.ImageFormat.Png);

            int i = 0;
            //Image imgTemp = Metafile.FromStream(new MemoryStream(range.EnhMetaFileBits));

            //imgTemp.Save(@"C:\Users\Word\Desktop\1233.png", System.Drawing.Imaging.ImageFormat.Png);
        }

        #region 得到光标在屏幕上的位置
        [DllImport("user32")]
        private static extern bool GetCursorPos(out System.Windows.Point lpPoint);


        private System.Windows.Point CaretPos()
        {
            System.Windows.Point showPoint = new System.Windows.Point();
            GetCursorPos(out showPoint);
            return showPoint;
        }
        #endregion


        private Microsoft.Office.Interop.Word.Application WordApp;
        private Microsoft.Office.Interop.Word.Document WordDoc;

        private void button7_Click_1(object sender, EventArgs e)
        {
            //System.Drawing.Point buttomPos = GetPositionForShowing(Globals.ThisAddIn.Application.Selection);
            //this.button7.Text = "selectPos"+buttomPos.X+":"+buttomPos.Y;

            WordApp = new Microsoft.Office.Interop.Word.Application();

            object filename = @"D:\BaoProject\高分宝三期\2014届高三数学理科模拟试题（定稿）.doc";
            Object Nothing = System.Reflection.Missing.Value;
            WordDoc = WordApp.Documents.Add(ref filename, ref Nothing, ref Nothing, ref Nothing);

            for (int i = 0; i < rangeList.Count;i++ )
            {
                string fileName = GetBetterImgFinal(rangeList[i].RangeStart,rangeList[i].RangeEnd);

                rangeList[i].Question = fileName;
            }
            jsonFileHelper.WriteFileString(JsonConvert.SerializeObject(rangeList), @"D:\新建文件夹\imgJsonList.json");
            Clipboard.Clear();
            //避免弹出normal.dotm被使用的对话框,自动保存模板
            WordApp.NormalTemplate.Saved = true;

            //先关闭打开的文档（注意saveChanges选项）
            Object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            Object originalFormat = Type.Missing;
            Object routeDocument = Type.Missing;
            WordApp.Documents.Close(ref saveChanges, ref originalFormat, ref routeDocument);
            
            //若已经没有文档存在，则关闭应用程序
            if (WordApp.Documents.Count == 0)
            {
                WordApp.Quit(Type.Missing, Type.Missing, Type.Missing);
            }

            new FormTest().Show();
        }

        private string GetBetterImgFinal(int rangeStart,int RangeEnd)
        {
            
            Range rangeFrom = WordDoc.Range(rangeStart, RangeEnd);
            rangeFrom.Copy();

            Range range = Globals.ThisAddIn.Application.ActiveDocument.Range();
            range.Font.Name = "宋体";
            range.Font.Size = (float)5;
            Globals.ThisAddIn.Application.ActiveDocument.Range().Text = "                                                                                                                                                                                               ";
            Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(range);


            range.SetRange(0, 0);
            range.Select();
            System.Drawing.Point leftTopPos = GetPositionForShowing(Globals.ThisAddIn.Application.Selection);

            range = Globals.ThisAddIn.Application.ActiveDocument.Range();
            range.SetRange(range.End, range.End);
            range.Select();
            System.Drawing.Point rightTopPos = GetPositionForShowing(Globals.ThisAddIn.Application.Selection);

            range.Text = "\n";
            range = Globals.ThisAddIn.Application.ActiveDocument.Range();
            range.SetRange(range.End, range.End);
            range.Select();
            range.Paste();

            

            range = Globals.ThisAddIn.Application.ActiveDocument.Range();
            range.SetRange(range.End, range.End);
            range.Font.Name = "宋体";
            range.Font.Size = (float)5;
            range.Text = "\n  ";
            
            range.SetRange(0, 0);
            Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(range);

            range = Globals.ThisAddIn.Application.ActiveDocument.Range();
            range.SetRange(range.End-1, range.End);
            range.Select();

            System.Threading.Thread.Sleep(1000);
            System.Drawing.Point buttomPos = GetPositionForShowing(Globals.ThisAddIn.Application.Selection);

            range=Globals.ThisAddIn.Application.ActiveDocument.Range();
            range.GrammarChecked = false;
            range.SpellingChecked = false;
            
            System.Threading.Thread.Sleep(1000);
            Bitmap image = new Bitmap(rightTopPos.X - leftTopPos.X+10, buttomPos.Y - leftTopPos.Y-20);
            Graphics g = Graphics.FromImage(image);
            g.CopyFromScreen(leftTopPos.X, leftTopPos.Y, 0, 0, image.Size);

            string fileName = System.Guid.NewGuid().ToString();
            image.Save(@"D:\新建文件夹\" +fileName + ".png", System.Drawing.Imaging.ImageFormat.Png);

            

            this.button7.Text = "leftTop" + leftTopPos.X + ":" + leftTopPos.Y + "rightTop" + rightTopPos.X + ":" + rightTopPos.Y + "Buttom" + buttomPos.X + ":" + buttomPos.Y;

            //WordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            //WordApp.Quit(ref Nothing, ref Nothing, ref Nothing);

            return fileName + ".png";

        }

        private static System.Drawing.Point GetPositionForShowing(Microsoft.Office.Interop.Word.Selection Sel)
        {
            // get range postion
            int left = 0;
            int top = 0;
            int width = 0;
            int height = 0;
            Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.GetPoint(out left, out top, out width, out height, Sel.Range);

            System.Drawing.Point currentPos = new System.Drawing.Point(left, top);
            //if (Screen.PrimaryScreen.Bounds.Height - top > 340)
            //{
            //    currentPos.Y += 20;
            //}
            //else
            //{
            //    currentPos.Y -= 320;
            //}
            return currentPos;
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            Range range = Globals.ThisAddIn.Application.Selection.Range;
            //System.Drawing.Point buttomPos = GetPositionForShowing(Globals.ThisAddIn.Application.Selection);

            //this.button8.Text = "POS" + buttomPos.X + ":" + buttomPos.Y + "range" + range.Start + ":" + range.End;

            

            //Sentences sentences = range.Sentences;
            //string sentenceSE = "";
            //for (int i = 0; i < sentences.Count; i++)
            //{
            //    Range sentenceRange = sentences[i + 1];
            //    sentenceSE = sentenceSE + "\n" + sentenceRange.Start + ":" + sentenceRange.End;
            //}
            //MessageBox.Show("Sentences:\n" + sentenceSE);

            Paragraphs paragraphs = range.Paragraphs;
            string paragraphSE = "";
            for (int i = 0; i < paragraphs.Count; i++)
            {
                Paragraph paragraph = paragraphs[i + 1];
                paragraphSE = paragraphSE + "\n" + paragraph.Range.Start + ":" + paragraph.Range.End;
                
            }

            Tables tables = range.TopLevelTables;
            for (int i = 0; i < tables.Count; i++)
            {
                Table table = tables[1 + i];
                paragraphSE = paragraphSE + "Table" + table.Range.Start + ":" + table.Range.End;
            }

            MessageBox.Show("Paragraphs" + paragraphSE);
            //Cells cells = range.Cells;
            //for(int i=0;i<cells.Count;i++)
            //{
            //    Cell cell=cells[i+1];
            //    MessageBox.Show(cell.Range.Text);
            //}

            
            
        }

    }
}
