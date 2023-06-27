using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace programma
{
    public partial class Form1 : Form
    {
        Word.Application wordApp;
        string[] marks;
        List<string> namefiles;
        string exePath = AppDomain.CurrentDomain.BaseDirectory;
        int m = 0;
        Word.Document doc;
        public Form1()
        {        
            InitializeComponent();
            dataGridView1.RowHeadersVisible = false;
            dataGridView2.RowHeadersVisible = false;
        }

        private void ExportButton_Click_1(object sender, EventArgs e)
        {
             try {
                wordApp = new Word.Application();
                button3.Visible = false;
                label3.Visible = false;
                m = 0;
                dataGridView1.Rows.Clear();
                for (int j = 0; j < 2; j++)
                {
                    if (dataGridView2.SelectedCells[j].Selected == true)
                    {
                        m = j;
                        break;
                    }
                }
                int i = 0;
                object path1 = Path.Combine(exePath, ("shablon\\" + dataGridView2.SelectedCells[m].Value.ToString() + ".docx"));
                var path2 = Path.Combine(exePath, "docum\\" + dataGridView2.SelectedCells[m].Value.ToString() + ".docx");
                object nullobj = System.Reflection.Missing.Value;
                doc = wordApp.Documents.Open(
                    ref path1, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj);
                doc.ActiveWindow.Selection.WholeStory();
                doc.ActiveWindow.Selection.Copy();
                IDataObject data = Clipboard.GetDataObject();
                string Content = data.GetData(DataFormats.Text).ToString();
                wordApp.Visible = false;
                marks = new string[Regex.Matches(Content, @"{[[a-zA-Z0-9_]*}").Count];
                foreach (Match match in Regex.Matches(Content, @"{[[a-zA-Z0-9_]*}"))
                {

                    dataGridView1.Rows.Add(match.Value);

                    marks[i] = match.Value;
                    i++;
                }
                for (int n=0; n < dataGridView1.Rows.Count; n++)
                {
                    dataGridView1.Rows[n].Cells["Column2"].Value = null;
                }
                   wordApp.Quit();
            }
           catch
            {
                MessageBox.Show("Произошла ошибка");
            }
            
        }

        private void ReplaceWord(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }


        private void button1_Click(object sender, EventArgs e)
        {
            try {
                int t = 0;
                wordApp = new Word.Application();
                var path1 = Path.Combine(exePath, "shablon\\" + dataGridView2.SelectedCells[m].Value.ToString() + ".docx");
                var path2 = Path.Combine(exePath, "docum\\" + dataGridView2.SelectedCells[m].Value.ToString() + ".docx");
                var wordDocument = wordApp.Documents.Open(path1);
                wordApp.Visible = false;
                var range = wordDocument.Content;
                range.Find.ClearFormatting();
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells["Column2"].Value != null)
                    {
                        ReplaceWord(marks[i], dataGridView1.Rows[i].Cells["Column2"].Value.ToString(), wordDocument);
                        t++;
                    }
                    else break;
                }
                if (t == dataGridView1.Rows.Count)
                {
                    wordDocument.SaveAs2(path2);
                    button3.Visible = true;
                    label3.Visible = true;
                    doc.Close();
                    wordApp.Quit();
                }
                else
                {
                    MessageBox.Show("Произошла ошибка\n  Введите данные");
                    wordApp.Quit();
                }
            }
            catch
            {
                for (int n = 0; n < dataGridView1.Rows.Count; n++)
                {
                    dataGridView1.Rows[n].Cells["Column2"].Value = null;
                }
                wordApp.Quit();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
            var path1 = Path.Combine(exePath, "shablon\\");
            int i = 0;
            var dir = new DirectoryInfo(path1);// папка с файлами 
            dataGridView2.Columns.Add("SpisokShablon", "Список шаблонов");
            namefiles = new List<string>();
            foreach (FileInfo file in dir.GetFiles()) // извлекаем все файлы и кидаем их в список 
            {
                namefiles.Add(Path.GetFileNameWithoutExtension(file.FullName)); // получаем полный путь к файлу и потом вычищаем ненужное, оставляем только имя файла. 
                dataGridView2.Rows.Add(namefiles[i]);
                i++;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            var path2 = Path.Combine(exePath, "docum\\" + dataGridView2.SelectedCells[m].Value.ToString() + ".docx");
            System.Diagnostics.Process.Start(path2);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var path2 = Path.Combine(exePath, "shablon\\");
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = path2;
            openFileDialog1.Filter = "docx files (*.docx)|*.docx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filename = openFileDialog1.FileName;
                System.Diagnostics.Process.Start(filename);
            }

        }
    }
}
