using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UAZik
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            textBox1.Text = Form1.Get_list();
            textBox2.Text = Form1.Get_row();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1.Set_list(textBox1.Text);
            Form1.Set_row(textBox2.Text);

            #region Создание файла и запись данных в него
            //Создаем переменную для создания/удаления файла
            FileInfo fileInf = new FileInfo("Settings.ini");

            //Создаем переменную для записи в файл и создадем файл на жестком диске
            StreamWriter fileWrite = new StreamWriter(fileInf.Create());

            //Записываем построчно данные в файл
            fileWrite.WriteLine(textBox1.Text);
            fileWrite.WriteLine(textBox2.Text);

            //Сохраняем данные и закрываем файл
            fileWrite.Close();
            #endregion
            Close();
        }
    }
}
