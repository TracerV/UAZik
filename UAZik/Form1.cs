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
using Excel = Microsoft.Office.Interop.Excel;

namespace UAZik
{
    public partial class Form1 : Form
    {
        string fileName1 = "C:\\Users\\Tracer\\Desktop\\УАЗик\\Отслеживание заявок дилеров.xlsx";
        string fileName2 = "C:\\Users\\Tracer\\Desktop\\УАЗик\\удалять.xlsx";
        int count_toDel;
        int count_toDel1;
        private static string list;
        private static string row;
        int count1,count2;
        public Form1()
        {
            InitializeComponent();
            #region Читаем данные из файла настроек
            //Обрабатываем возможную ошибку
            try
            {
                #region Чтение данных из файла Setting.ini
                //Объявляеим переменную для чтения из файла и привязываем ее к фалу
                StreamReader fileRead = new StreamReader("Settings.ini");

                try
                {
                    //Считываем данные из файла
                    list = fileRead.ReadLine();
                    row = fileRead.ReadLine();
                }
                catch
                {
                    list = "2";
                    row = "2";
                }
                finally
                {
                    //Закрываем файл
                    fileRead.Close();
                }
                #endregion
            }
            catch //Код если произошла ошибка
            {
                //Помещаем значения по умолчанию
                list = "2";
                row = "2";
                #region Создание файла и запись данных в него
                //Создаем переменную для создания/удаления файла
                FileInfo fileInf = new FileInfo("Settings.ini");

                //Создаем переменную для записи в файл и создадем файл на жестком диске
                StreamWriter fileWrite = new StreamWriter(fileInf.Create());

                //Записываем построчно данные в файл
                fileWrite.WriteLine(list);
                fileWrite.WriteLine(row);

                //Сохраняем данные и закрываем файл
                fileWrite.Close();
                #endregion
            }
            #endregion
        }

        #region Get/Set
        public static string Get_list()
        {
            return list;
        }

        public static string Get_row()
        {
            return row;
        }

        public static void Set_list(string data)
        {
            list = data;
        }

        public static void Set_row(string data)
        {
            row = data;
        }
        #endregion

        private void clean_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp1 = new Excel.Application(); 
            Excel.Application xlApp2 = new Excel.Application(); 
            Excel.Workbook xlWB1, xlWB2; //рабочая книга            
            Excel.Worksheet xlSht1, xlSht2; //лиcn Ecxel 
            fileName1 = textBox1.Text;
            fileName2 = textBox2.Text;
            count_toDel = 0;

            #region Получение массива строк,которые нужно удалить

            xlWB2 = xlApp2.Workbooks.Open(fileName2); //название файла Excel
            xlSht2 = xlWB2.Sheets[Convert.ToInt32(list)]; //название листа или можно так если лист первый по порядку - xlWB.Sheets[1];
            int iLastRow2 = xlSht2.Cells[xlSht2.Rows.Count, "B"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце B 

            for (int i = 1; i <= iLastRow2; i++)
            {
                //Получаем значение из ячейки и преобразуем в строку и записываем в массив toDel
                if (xlSht2.Cells[i, Convert.ToInt32(row)].Value != null)
                {
                    count_toDel++;
                }
                else
                { }
            }

            string[] toDel = new string[count_toDel]; //выделяем память для массива на удаление
            count_toDel = 0;

            for (int i=1;i<= iLastRow2; i++)
            {
                if (xlSht2.Cells[i, Convert.ToInt32(row)].Value != null)
                {
                    Excel.Range forYach = xlSht2.Cells[i, Convert.ToInt32(row)] as Excel.Range;
                    string yach = forYach.Value2.ToString();
                    toDel[count_toDel] = yach;
                    count_toDel++;
                }
                else { }
            }
            #endregion

            #region Чистка исходного файла от лишних ячеек

            xlWB1 = xlApp1.Workbooks.Open(fileName1); //название файла Excel
            xlSht1 = xlWB1.Sheets[1]; //название листа или можно так если лист первый по порядку - xlWB.Sheets[1];
            int iLastRow1 = xlSht1.Cells[xlSht1.Rows.Count, "B"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце B

            string[] arrayOld = new string[iLastRow1-5]; //выделяем память для массива данных

            count_toDel1 = 0;

            for (int i=5;i<iLastRow1;i++)
            {
                if (xlSht1.Cells[i, Convert.ToInt32(row)].Value != null)
                {
                    Excel.Range forYach = xlSht1.Cells[i, 8] as Excel.Range;
                    string yach = forYach.Value2.ToString();
                    arrayOld[count_toDel1] = yach;
                    count_toDel1++;
                }
                else { }
            }
            count1 = 0;
            for (int i = 5; i < iLastRow1; i++)
            {
                for (int j = 5; j < iLastRow1; j++)
                {
                    if (i!=j)
                    {
                        if ((arrayOld[j - 5].Contains(arrayOld[i - 5])) && (arrayOld[i - 5] != arrayOld[j - 5])&&(Convert.ToInt32(arrayOld[i - 5].Length)<11))
                        {
                            xlSht1.Rows[i-count1].Delete();
                            count1++;
                            break;
                        } else { }
                    } else { }
                }
            }
            #endregion

            string[] arrayNew = new string[iLastRow1 - 5-count1]; //выделяем память для нового массива данных

            for (int i = 5; i < iLastRow1-count1; i++)
            {
                    Excel.Range forYach = xlSht1.Cells[i, 8] as Excel.Range;
                    string yach = forYach.Value2.ToString();
                    arrayNew[i-5] = yach;
            }

            #region Чистим редактированный файл из массива toDel
            count2 = 0;
            for (int i = 0; i < arrayNew.Length; i++)
            {
                for (int j = 0; j < toDel.Length; j++)
                {
                    if (toDel[j] == arrayNew[i])
                    {
                        xlSht1.Rows[i+5-count2].Delete();
                        count2++;
                    }
                }
            }
            #endregion

            xlWB1.Close(true);
            xlWB2.Close(false);


            MessageBox.Show("Удалено: " + count1 + " + " + count2 + " строк");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)

            {
                textBox1.Text = openFileDialog1.FileName;
             //   fileName1 = textBox1.Text;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)

            {
                textBox2.Text = openFileDialog1.FileName;
             //   fileName2 = textBox2.Text;
            }
        }

        private void Settings_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }
    }
}
