using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SelectFromBD
{
    public partial class Form1 : Form
    {
        public static string connectString;
        private OleDbConnection myConnection;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            string BDsource = textBox5.Text;
            connectString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + BDsource;
            myConnection = new OleDbConnection(connectString);
            try
            {
                myConnection.Open();
                MessageBox.Show("Подключение выполнено успешно!", "Внимание!");
            }
            catch
            {
                MessageBox.Show("Ошибка подключения!", "Внимание!");
                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string TableName = textBox2.Text;
                int counter = 1;

                OleDbCommand command1 = new OleDbCommand();
                command1.Connection = myConnection;

                string query1 = "SELECT Номер_поставщика, Товар, Цена FROM " + TableName + " ORDER BY Номер_поставщика, Цена;";
                command1.CommandText = query1;

                OleDbDataAdapter da1 = new OleDbDataAdapter(command1);
                DataTable dt1 = new DataTable();
                da1.Fill(dt1);
                dataGridView1.DataSource = dt1;

                counter++;
            }
            catch
            {
                MessageBox.Show("Ошибка поиска полей из таблицы!\nУбедитесь что поля названы верно.(прочтите readme.txt)\nЛибо проверьте правильность введения названия таблицы.", "Внимание!");
            }

            myConnection.Close();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        OpenFileDialog ofd = new OpenFileDialog();

        private void button3_Click(object sender, EventArgs e)
        {
            if(ofd.ShowDialog() == DialogResult.OK)
            {
                textBox5.Text = ofd.FileName;
            }
        }
    }
}
