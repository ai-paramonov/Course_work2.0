using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace MyAPP
{
    
    public partial class Form1 : Form
    {
        public static string connectString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = CompanyDB.mdb";
        private OleDbConnection myConnection;
        public Form1()
        {
            InitializeComponent();
            myConnection = new OleDbConnection(connectString);
            myConnection.Open();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "companyDBDataSet.Сотрудники". При необходимости она может быть перемещена или удалена.
            this.сотрудникиTableAdapter.Fill(this.companyDBDataSet.Сотрудники);

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            myConnection.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int kod = Convert.ToInt32(textBox1.Text);
            string query = "DELETE FROM Сотрудники WHERE [Код сотрудника] = " + kod;
            OleDbCommand command = new OleDbCommand(query, myConnection);
            command.ExecuteNonQuery();
            MessageBox.Show("Данные о сотруднике удалены");
            this.сотрудникиTableAdapter.Fill(this.companyDBDataSet.Сотрудники);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int kod = Convert.ToInt32(textBox2.Text);
            string query = "UPDATE Сотрудники SET Должность ='" + textBox3.Text + "' WHERE [Код сотрудника] = " + kod;
            OleDbCommand command = new OleDbCommand(query, myConnection);
            command.ExecuteNonQuery();
            MessageBox.Show("Должность обновлена");
            this.сотрудникиTableAdapter.Fill(this.companyDBDataSet.Сотрудники);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Owner = this;
            f2.Show();
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.сотрудникиTableAdapter.Fill(this.companyDBDataSet.Сотрудники);
        }

        private void сотрудникиПоДолжностиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Owner = this;
            f3.Show();
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            this.сотрудникиTableAdapter.Fill(this.companyDBDataSet.Сотрудники);
        }
    }
}
