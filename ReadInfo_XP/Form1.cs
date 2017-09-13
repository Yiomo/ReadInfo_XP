using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System;
using System.Diagnostics;

namespace ReadInfo_XP
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            FolderBrowserDialog openfolder = new FolderBrowserDialog();
            DialogResult result = openfolder.ShowDialog();
            if (result == DialogResult.Cancel)
            {
                return;
            }
            //textBlock.Text = openfolder.SelectedPath.Trim();
            DirectoryInfo TheFolder = new DirectoryInfo(openfolder.SelectedPath);
            foreach (FileInfo NextFile in TheFolder.GetFiles())
            {
                if (NextFile.Name.Contains(".txt") || NextFile.Name.Contains(".TXT"))
                {
                    StreamReader readfile = new StreamReader(openfolder.SelectedPath + "\\" + NextFile.Name);
                    listBox1.Items.Add(NextFile.Name);
                }
            }
            int a = listBox1.Items.Count;//统计多少文件

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            app.UserControl = true;
            Workbook workbook;
            Worksheet worksheet=null;
            workbook = app.Workbooks.Add( );
            worksheet = workbook.Worksheets.Add(System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value)as Worksheet ;
            worksheet.Name = "total";
            for (int i = 0; i < a; i++)//逐行写入
            {
                //textBlock1.Text = "";
                //textBlock.Text = a.ToString();
                string[] lines;
                lines = File.ReadAllLines(openfolder.SelectedPath + "\\" + listBox1.Items[i].ToString());
                for (int j = 0; j < lines.Length; j++)
                {
                    //textBlock1.Text = textBlock1.Text + lines[j] + "\n";
                    //ws.cells[i + 1, j + 1] = dv[i][j].tostring();
                    worksheet.Cells[i + 1, j + 1] = lines[j];
                }
            }
            workbook.SaveAs(openfolder.SelectedPath + "\\" + "sum.xls", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            workbook.Close(System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            MessageBox.Show("Done.");

            Process[] process = Process.GetProcessesByName("excel");
            foreach (Process p in process)
            {
                if (!p.HasExited)  // 如果程序没有关闭，结束程序
                {
                    p.Kill();
                    p.WaitForExit();
                }
            }
        }
    }
}
