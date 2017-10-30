using CheckOnWorkTool.service;
using CheckOnWorkTool.util;
using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace CheckOnWorkTool
{
    public partial class main : Form
    {
        public main()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void process1_Exited(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }


        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = @"Excel2007文件|*.xlsx|Excel2003文件|*.xls";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;
                if (file == "")
                {
                    return;
                }
                this.textBox1.Text = file;
                MessagesUtil.addMsg(Properties.Resources.writeExcelLb + file);
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog3.Filter = @"Excel2007文件|*.xlsx|Excel2003文件|*.xls";
            DialogResult result = openFileDialog3.ShowDialog();
            if (result == DialogResult.OK)
            {
                string file = openFileDialog3.FileName;
                if (file == "")
                {
                    return;
                }
                this.textBox3.Text = file;
                MessagesUtil.addMsg(Properties.Resources.monthlyLb + file);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //流数量
            CheckOnWorkService.currentThreadNum += 1;
            //开始处理
            new CheckOnWorkService().writeMegToExcel(textBox1.Text, textBox3.Text);
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.Text.Length > 0)
            {
                textBox4.Focus();
                textBox4.Select(textBox4.TextLength, 0);
                textBox4.ScrollToCaret();

            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
