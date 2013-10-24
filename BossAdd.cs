using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Data.OleDb;
using System.Data.Common;
using System.Collections;
using System.IO;

namespace SPA
{
    public partial class BossAdd : Form
    {
        string s = " ";
        bool a;


        List<string> list = new List<string>();


        OleDbConnection myOleDbConnection;
        OleDbDataAdapter myDataAdapter;
        DataSet myDataSet;
        public OleDbConnection obj_connect = null;
        string connectionString;
        bool SPA_ = false;
        public BossAdd()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string cmd = "INSERT INTO Персонал (Фамилия,Имя,Отчество,Специальность)  VALUES ('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "', '" + comboBox1.Text + "' )";

                myDataAdapter.InsertCommand = new OleDbCommand(cmd, myOleDbConnection);

                myDataAdapter.InsertCommand.Connection.Open();
                myDataAdapter.InsertCommand.ExecuteNonQuery();
                
                myDataAdapter.InsertCommand.Connection.Close();

                myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM Персонал", myOleDbConnection);
                myDataAdapter.SelectCommand.Connection.Open();
                myDataAdapter.SelectCommand.ExecuteNonQuery();
                myDataAdapter.SelectCommand.Connection.Close();
                MessageBox.Show("Новый сотрудник назначен");
                myDataSet.Tables["Персонал"].Clear();
                myDataAdapter.Fill(myDataSet, "Персонал");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                obj_connect = null;
            }
        }

        private void BossAdd_Load(object sender, EventArgs e)
        {

            ToolTip t = new ToolTip();

            t.SetToolTip(this.button15, "Раскрыть список");
            t.SetToolTip(this.button18, "Свернуть список");

            t.SetToolTip(this.button1, "Добавить специалиста");
            t.SetToolTip(this.button7, "Удалить специалиста");

            t.SetToolTip(this.button3, "Добавить специальность");
            t.SetToolTip(this.button8, "Удалить специальность");

            t.SetToolTip(this.button5, "Добавить процедуру");
            t.SetToolTip(this.button9, "Удалить процедуру");
            t.SetToolTip(this.button10, "Добавить описание к процедуре");


            t.SetToolTip(this.button14, "Добавить Spa программу");
            t.SetToolTip(this.button12, "Удалить Spa программу");


            connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + "data source=spa.mdb";

            myOleDbConnection = new OleDbConnection(connectionString);

            myOleDbConnection = new OleDbConnection(connectionString);
            myDataAdapter = new System.Data.OleDb.OleDbDataAdapter("SELECT * FROM Персонал", myOleDbConnection);
            myDataSet = new DataSet("Персонал");

            myDataAdapter.Fill(myDataSet, "Персонал");
            myDataAdapter.SelectCommand.Connection.Close();
            myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Клиенты", myOleDbConnection);
            myDataAdapter.SelectCommand.Connection.Open();
            myDataAdapter.SelectCommand.ExecuteNonQuery();
            myDataAdapter.Fill(myDataSet, "Клиенты");

            myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Время", myOleDbConnection);
            myDataAdapter.SelectCommand.ExecuteNonQuery();
            myDataAdapter.Fill(myDataSet, "Время");


            myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Процедуры", myOleDbConnection);
            myDataAdapter.SelectCommand.ExecuteNonQuery();
            myDataAdapter.Fill(myDataSet, "Процедуры");

            myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Специальности", myOleDbConnection);
            myDataAdapter.SelectCommand.ExecuteNonQuery();
            myDataAdapter.Fill(myDataSet, "Специальности");

            myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание", myOleDbConnection);
            myDataAdapter.SelectCommand.ExecuteNonQuery();
            myDataAdapter.Fill(myDataSet, "Расписание");


            myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM spa_процедуры", myOleDbConnection);
            myDataAdapter.SelectCommand.ExecuteNonQuery();
            myDataAdapter.Fill(myDataSet, "spa_процедуры");

            myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Персонал WHERE (Специальность = 'СПА-мастер')", myOleDbConnection);
            myDataAdapter.SelectCommand.ExecuteNonQuery();
            myDataAdapter.Fill(myDataSet, "SPA_Персонал");

            myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Специальности WHERE (Not Название='СПА-мастер')", myOleDbConnection);
            myDataAdapter.SelectCommand.ExecuteNonQuery();
            myDataAdapter.Fill(myDataSet, "NonSPA_Специальности");


            myDataAdapter.SelectCommand.Connection.Close();

            this.dataGridView3.DataSource = myDataSet.Tables[0];
            this.dataGridView1.DataSource = myDataSet.Tables["Специальности"].DefaultView;
            this.dataGridView2.DataSource = myDataSet.Tables["Процедуры"].DefaultView;
            this.dataGridView4.DataSource = myDataSet.Tables["spa_процедуры"].DefaultView;

            this.dataGridView1.Columns["ID_Специальности"].Visible = false;
            this.dataGridView2.Columns["ID_Процедуры"].Visible = false;
            this.dataGridView2.Columns["ID_Специальности"].Visible = false;
            this.dataGridView2.Columns["Описание"].Visible = false;
            this.dataGridView3.Columns["ID_Персонала"].Visible = false;
            this.dataGridView4.Columns["ID"].Visible = false;
            this.dataGridView4.Columns["ID_Персонала"].Visible = false;

            comboBox1.DataSource = myDataSet.Tables["Специальности"].DefaultView;
            comboBox1.DisplayMember = "Название";

            comboBox2.DataSource = myDataSet.Tables["SPA_Персонал"].DefaultView;
            comboBox2.DisplayMember = "Фамилия";

            comboBox3.DataSource = myDataSet.Tables["NonSPA_Специальности"].DefaultView;
            comboBox3.DisplayMember = "Название";


            checkedListBox1.DataSource = myDataSet.Tables["Процедуры"].DefaultView;
            checkedListBox1.DisplayMember = "Название";


        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string cmd = "INSERT INTO Специальности (Название)  VALUES ('" + textBox4.Text + "')";
            try
            {
                myDataAdapter.InsertCommand = new OleDbCommand(cmd, myOleDbConnection);
                myDataAdapter.InsertCommand.Connection.Open();
                myDataAdapter.InsertCommand.ExecuteNonQuery();
                myDataAdapter.InsertCommand.Connection.Close();

                myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM Специальности", myOleDbConnection);
                myDataAdapter.SelectCommand.Connection.Open();
                myDataAdapter.SelectCommand.ExecuteNonQuery();
                myDataAdapter.SelectCommand.Connection.Close();
                textBox4.Clear();

                myDataSet.Tables["Специальности"].Clear();
                myDataAdapter.Fill(myDataSet, "Специальности");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                obj_connect = null;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            myDataAdapter.DeleteCommand = new OleDbCommand("DELETE FROM Персонал WHERE ID_Персонала=" + dataGridView3.SelectedRows[0].Cells[0].Value, myOleDbConnection);
            try
            {
                myDataAdapter.DeleteCommand.Connection.Open();
                myDataAdapter.DeleteCommand.ExecuteNonQuery();
                myDataAdapter.DeleteCommand.Connection.Close();

                myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM Персонал", myOleDbConnection);
                myDataAdapter.SelectCommand.Connection.Open();
                myDataAdapter.SelectCommand.ExecuteNonQuery();
                myDataAdapter.SelectCommand.Connection.Close();

                myDataSet.Tables["Персонал"].Clear();
                myDataAdapter.Fill(myDataSet, "Персонал");
                MessageBox.Show("сотрудник освобожден от занимаеой должности и удален из базы");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                obj_connect = null;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {

            myDataAdapter.DeleteCommand = new OleDbCommand("DELETE FROM Специальности WHERE Название ='" + dataGridView1.SelectedRows[0].Cells[0].Value + "'", myOleDbConnection);
            try
            {
                myDataAdapter.DeleteCommand.Connection.Open();
                myDataAdapter.DeleteCommand.ExecuteNonQuery();
               
                myDataAdapter.DeleteCommand.Connection.Close();

                myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM Специальности", myOleDbConnection);
                myDataAdapter.SelectCommand.Connection.Open();
                myDataAdapter.SelectCommand.ExecuteNonQuery();
                myDataAdapter.SelectCommand.Connection.Close();

                myDataSet.Tables["Специальности"].Clear();
                myDataAdapter.Fill(myDataSet, "Специальности");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                obj_connect = null;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int a = 0;
            string cc = "SELECT Специальности.ID_Специальности FROM Специальности WHERE (((Специальности.Название)='" + comboBox3.Text + "'));";
            OleDbConnection myConn = new OleDbConnection(connectionString);
            myConn.Open();
            OleDbCommand cmd1 = new OleDbCommand(cc, myConn);
            a = Convert.ToInt32(cmd1.ExecuteScalar());


            string cmd = "INSERT INTO Процедуры (Название,Продолжительность,ID_Специальности)  VALUES ('" + textBox5.Text + "'," + textBox9.Text + "," + a + ")";
            try
            {
                myDataAdapter.InsertCommand = new OleDbCommand(cmd, myOleDbConnection);

                myDataAdapter.InsertCommand.Connection.Open();
                myDataAdapter.InsertCommand.ExecuteNonQuery();
                myDataAdapter.InsertCommand.Connection.Close();

                myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM Процедуры", myOleDbConnection);
                myDataAdapter.SelectCommand.Connection.Open();
                myDataAdapter.SelectCommand.ExecuteNonQuery();
                myDataAdapter.SelectCommand.Connection.Close();
                textBox5.Clear();

                myDataSet.Tables["Процедуры"].Clear();
                myDataAdapter.Fill(myDataSet, "Процедуры");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                obj_connect = null;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            myDataAdapter.DeleteCommand = new OleDbCommand("DELETE FROM Процедуры WHERE ID_Процедуры=" + dataGridView2.SelectedRows[0].Cells[0].Value + "", myOleDbConnection);
            try
            {
                myDataAdapter.DeleteCommand.Connection.Open();
                myDataAdapter.DeleteCommand.ExecuteNonQuery();
                myDataAdapter.DeleteCommand.Connection.Close();

                myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM Процедуры", myOleDbConnection);
                myDataAdapter.SelectCommand.Connection.Open();
                myDataAdapter.SelectCommand.ExecuteNonQuery();
                myDataAdapter.SelectCommand.Connection.Close();

                myDataSet.Tables["Процедуры"].Clear();
                myDataAdapter.Fill(myDataSet, "Процедуры");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                obj_connect = null;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //myDataAdapter.UpdateCommand = new OleDbCommand("UPDATE Процедуры SET [Описание] = '" + textBox6.Text + "' WHERE [Название] ='" + dataGridView2.SelectedRows[0].Cells[0].Value + "'", myOleDbConnection);
            //try
            //{
            //    myDataAdapter.UpdateCommand.Connection.Open();
            //    myDataAdapter.UpdateCommand.ExecuteNonQuery();
            //    myDataAdapter.UpdateCommand.Connection.Close();
            //    textBox6.Clear();

            //    myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM Процедуры", myOleDbConnection);
            //    myDataAdapter.SelectCommand.Connection.Open();
            //    myDataAdapter.SelectCommand.ExecuteNonQuery();
            //    myDataAdapter.SelectCommand.Connection.Close();

            //    myDataSet.Tables["Процедуры"].Clear();
            //    myDataAdapter.Fill(myDataSet, "Процедуры");
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //    obj_connect = null;
            //}
        }


        private void button14_Click(object sender, EventArgs e)
        {

            int a = 0;
            string cc = "SELECT Персонал.ID_Персонала FROM Персонал WHERE (((Персонал.Фамилия)='" + comboBox2.Text + "'));";
            OleDbConnection myConn = new OleDbConnection(connectionString);
            myConn.Open();
            OleDbCommand cmd1 = new OleDbCommand(cc, myConn);
            a = Convert.ToInt32(cmd1.ExecuteScalar());



            checkedListBox1.ClearSelected();
            checkedListBox1.Height = 20;
            int i = 0;
            for (i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }


            checkedListBox1.Height = 20;
            List<string> uniqueList = new List<string>(list.Distinct());
            int k = 0;
            foreach (var item in uniqueList)
            {
                s += uniqueList[k];
                k++;
            }
            i = 0;

            uniqueList.Remove(" <> ");

            list.Clear();


            string cmd = "INSERT INTO spa_процедуры (Название,Описание,Продолжительность,ID_Персонала)  VALUES ('" + textBox8.Text + "','" + s + "'," + textBox7.Text + "," + a + ")";
            //string cmd = String.Format("INSERT INTO spa_процедуры (Описание)  VALUES ('{0}')", s);
            try
            {
                myDataAdapter.InsertCommand = new OleDbCommand(cmd, myOleDbConnection);

                myDataAdapter.InsertCommand.Connection.Open();
                myDataAdapter.InsertCommand.ExecuteNonQuery();
                myDataAdapter.InsertCommand.Connection.Close();

                myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM spa_процедуры", myOleDbConnection);
                myDataAdapter.SelectCommand.Connection.Open();
                myDataAdapter.SelectCommand.ExecuteNonQuery();
                myDataAdapter.SelectCommand.Connection.Close();
                textBox5.Clear();

                myDataSet.Tables["spa_процедуры"].Clear();
                myDataAdapter.Fill(myDataSet, "spa_процедуры");
                s = null;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                obj_connect = null;
            }

            checkedListBox1.ClearSelected();

        }



        private void button15_Click(object sender, EventArgs e)
        {
            checkedListBox1.ClearSelected();
            checkedListBox1.Height = 100;
        }

        private void button15_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void tabControl1_Click(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_Click(object sender, EventArgs e)
        {
            checkedListBox1.Height = 100;
        }

        private void checkedListBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //checkedListBox1.Height = 20;
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {

            list.Add(" <" + checkedListBox1.Text + "> ");
            list.Remove(" <> ");
        }








        private void button18_Click(object sender, EventArgs e)
        {
            checkedListBox1.ClearSelected();
            checkedListBox1.Height = 20;
            int i = 0;
            for (i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }

        }

        private void button12_Click(object sender, EventArgs e)
        {
            myDataAdapter.DeleteCommand = new OleDbCommand("DELETE FROM spa_процедуры WHERE ID=" + dataGridView4.SelectedRows[0].Cells[0].Value + "", myOleDbConnection);
            try
            {
                myDataAdapter.DeleteCommand.Connection.Open();
                myDataAdapter.DeleteCommand.ExecuteNonQuery();
                myDataAdapter.DeleteCommand.Connection.Close();

                myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM spa_процедуры", myOleDbConnection);
                myDataAdapter.SelectCommand.Connection.Open();
                myDataAdapter.SelectCommand.ExecuteNonQuery();
                myDataAdapter.SelectCommand.Connection.Close();

                myDataSet.Tables["spa_процедуры"].Clear();
                myDataAdapter.Fill(myDataSet, "spa_процедуры");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                obj_connect = null;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
