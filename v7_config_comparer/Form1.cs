using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using v7MetaData;

namespace v7_config_comparer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog Fd = new OpenFileDialog();
            Fd.Filter = "Конфигурации 1C|*.md";



            Fd.Title = "Выберите первую конфигурацию";
            if (Fd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;

            Program.FirstFileName = Fd.FileName;

            Fd.Title = "Выберите вторую конфигурацию";
            if (Fd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;

            Program.SecondFileName = Fd.FileName;

            Program.First = new OleStorage.TaskItem(Program.FirstFileName);
            Program.Second = new OleStorage.TaskItem(Program.SecondFileName);

            textBox1.Text = Program.Second.CompareWith(Program.First, true, true).Replace("\n", "\r\n");
        }
    }
}
