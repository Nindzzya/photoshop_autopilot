using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApp
{
    public partial class Form1 : Form
    {
        Photoshop.Application pApplication;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button1.Enabled = true;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "正在获取实例，请稍候！由于程序启动需要时间，请耐心等待。";
            try
            {
                pApplication = new Photoshop.Application();
                textBox1.Text = "已获取实例，当前版本：" + pApplication.Version;
                button2.Enabled = true;
                button3.Enabled = true;
                button4.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
                button7.Enabled = true;
            }
            catch(Exception err)
            {
                MessageBox.Show(this, err.Message, "实例获取失败", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBox1.Text = "实例获取失败，错误详情如下：" + err.Message;
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "请选择图像文件";
            openFileDialog.Filter = "(*.jpg;*.png;*.bmp;*.jpeg;*.gif;*.tiff;*.ico)|*.jpg;*.png;*.bmp;*.jpeg;*.gif;*.tiff;*.ico";
            DialogResult dialogResult = openFileDialog.ShowDialog();
            if (!dialogResult.ToString().Equals("Cancel"))
            {
                try
                {
                    pApplication.Open(Path.GetFullPath(openFileDialog.FileName));
                }
                catch (Exception err)
                {
                    MessageBox.Show(this, err.Message, "运行出现异常", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                Photoshop.Document pDocument = pApplication.ActiveDocument;
                Photoshop.Layers pLayers = pDocument.Layers;
                string layersText = "";
                foreach (Photoshop.ArtLayer artLayer in pLayers)
                {
                    layersText += artLayer.Name + Environment.NewLine;
                }
                textBox1.Text = layersText;
            }
            catch (Exception err)
            {
                MessageBox.Show(this, err.Message, "运行出现异常", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                Photoshop.Document pDocument = pApplication.ActiveDocument;
                pDocument.ArtLayers.Add();
            }
            catch (Exception err)
            {
                MessageBox.Show(this, err.Message, "运行出现异常", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                long timestamp = (long)(DateTime.Now.ToUniversalTime() - new DateTime(1970, 1, 1)).TotalMilliseconds;
                pApplication.ActiveDocument.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\ps_autosave_" + timestamp.ToString() + ".jpg");
            }
            catch(Exception err)
            {
                MessageBox.Show(this, err.Message, "运行出现异常", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            pApplication.ActiveDocument.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                long timestamp = (long)(DateTime.Now.ToUniversalTime() - new DateTime(1970, 1, 1)).TotalMilliseconds;
                pApplication.ActiveDocument.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\ps_autosave_" + timestamp.ToString() + ".jpg");
                pApplication.Quit();
                Application.Restart();
            }
            catch(Exception err)
            {
                MessageBox.Show(this, err.Message, "运行出现异常", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                try
                {
                    pApplication.Quit();
                    Application.Restart();
                }
                catch
                {
                    return;
                }
                return;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                Photoshop.ArtLayer artLayer = pApplication.ActiveDocument.ActiveLayer;
                artLayer.FillOpacity = 50;
            }
            catch (Exception err)
            {
                MessageBox.Show(this, err.Message, "运行出现异常", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
    }
}
