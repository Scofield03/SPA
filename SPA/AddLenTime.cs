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

namespace SPA
{
  public partial class AddLenTime : Form
  {
      Form1 f1;
        OleDbConnection myOleDbConnection;
        OleDbDataAdapter myDataAdapter;
        DataSet myDataSet;
        public OleDbConnection obj_connect = null;
     
    public AddLenTime()
    {
      InitializeComponent();
      //this.Activate();
        string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + "data source=spa.mdb";//connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + "data source=C:\\Users\\Сева\\Desktop\\курсовая\\курсовая\\Мед_центр.mdb";
            myOleDbConnection = new OleDbConnection(connectionString);

            myOleDbConnection = new OleDbConnection(connectionString);
            myDataAdapter = new System.Data.OleDb.OleDbDataAdapter("SELECT * FROM Персонал", myOleDbConnection);
            myDataSet = new DataSet("Персонал");

            myDataAdapter.Fill(myDataSet, "Персонал");

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


            myDataAdapter.SelectCommand.Connection.Close();           

            this.dataGridView1.DataSource = myDataSet.Tables["Время"].DefaultView;
            this.dataGridView1.Columns["ID"].Visible = false;
            this.dataGridView1.Columns["Дата"].Visible = false;
            this.dataGridView1.Columns["flag"].Visible = false;

            comboBox1.DataSource = myDataSet.Tables["Процедуры"].DefaultView;
            comboBox1.DisplayMember = "Название";

            comboBox6.DataSource = myDataSet.Tables["Персонал"].DefaultView;
            comboBox6.DisplayMember = "Фамилия";

            comboBox7.DataSource = myDataSet.Tables["spa_процедуры"].DefaultView;
            comboBox7.DisplayMember = "Название";
    }

    private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
    {
      //
      //int a = 10, b = 10, c = 1990;
      //dateTimePicker1.Value = new DateTime(c, b, a);

      //dateTimePicker1.Format = DateTimePickerFormat.Custom;
      //dateTimePicker1.CustomFormat = "MM/dd/yyyy";
      //listBox1.Items.Add(dateTimePicker1.Value.ToShortDateString());
      //listBox1.Items.Add(dateTimePicker1.Value.Month.ToString());
     // MessageBox.Show(dateTimePicker1.Value.Date.ToString());
     

    }

    private void AddLenTime_Load(object sender, EventArgs e)
    {
      ToolTip t = new ToolTip();
      t.SetToolTip(this.button3, "Выйти");
      t.SetToolTip(this.button2, "Добавить");
      t.SetToolTip(this.button1, "Удалить");

      comboBox1.Text = null;
      comboBox7.Text = null;
      comboBox6.Text = null;
    }

      private void button1_Click(object sender, EventArgs e)
    {
      if (comboBox7.SelectedItem == null)
      {

        String s = dateTimePicker1.Value.ToShortDateString() + " " + comboBox2.SelectedItem.ToString() + ":" + comboBox3.SelectedItem.ToString() + " - " + comboBox4.SelectedItem.ToString() + ":" + comboBox5.SelectedItem.ToString() + "( " + comboBox1.Text+")";

       
          string cmd = "INSERT INTO Время (Дата,Специалист,Процедура,С,По,_Дата)  VALUES ('" + s + "','" + comboBox6.Text + "','" + comboBox1.Text + "','" + comboBox2.SelectedItem.ToString() + ":" + comboBox3.SelectedItem.ToString() + "','" + comboBox4.SelectedItem.ToString() + ":" + comboBox5.SelectedItem.ToString() + "','" + dateTimePicker1.Value.ToShortDateString() + Properties.Resources1.ResourceEntry;


          try
          {

            if (comboBox2.SelectedIndex == -1 || comboBox3.SelectedIndex == -1 || comboBox4.SelectedIndex == -1 || comboBox5.SelectedIndex == -1)
              MessageBox.Show("Не все поля времени заполнены!", "ОШИБКА!", MessageBoxButtons.OK, MessageBoxIcon.Error);


            if (Convert.ToInt32(comboBox4.SelectedItem.ToString()) > Convert.ToInt32(comboBox2.SelectedItem.ToString()) || ((Convert.ToInt32(comboBox4.SelectedItem.ToString()) == Convert.ToInt32(comboBox2.SelectedItem.ToString()) && Convert.ToInt32(comboBox5.SelectedItem.ToString()) > Convert.ToInt32(comboBox3.SelectedItem.ToString()))))
            {

              myDataAdapter.InsertCommand = new OleDbCommand(cmd, myOleDbConnection);

              myDataAdapter.InsertCommand.Connection.Open();
              myDataAdapter.InsertCommand.ExecuteNonQuery();
              myDataAdapter.InsertCommand.Connection.Close();

              myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM Время", myOleDbConnection);
              myDataAdapter.SelectCommand.Connection.Open();
              myDataAdapter.SelectCommand.ExecuteNonQuery();
              myDataAdapter.SelectCommand.Connection.Close();

              myDataSet.Tables["Время"].Clear();
              myDataAdapter.Fill(myDataSet, "Время");

            }
            else
            {
              MessageBox.Show("Введите корректное время!", "ОШИБКА!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
          }
          catch (Exception ex)
          {

            MessageBox.Show(ex.Message);
            obj_connect = null;

          }
        }
        else MessageBox.Show("трололошеньки", "Achtung");
        
      }
   

    private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
    {
       
            
    }

    private void button3_Click(object sender, EventArgs e)
    {
        this.Close();
       
    }

    private void label10_Click(object sender, EventArgs e)
    {

    }

    private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
    {

    }

    private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
    {

    }

    private void label2_Click(object sender, EventArgs e)
    {

    }

    private void AddLenTime_Activated(object sender, EventArgs e)
    {
       
    }

    private void button2_Click(object sender, EventArgs e)
    {
      myDataAdapter.DeleteCommand = new OleDbCommand("DELETE FROM Время WHERE ID= "+ dataGridView1.SelectedRows[0].Cells[0].Value, myOleDbConnection);
      try
      {
        myDataAdapter.DeleteCommand.Connection.Open();
        myDataAdapter.DeleteCommand.ExecuteNonQuery();
        myDataAdapter.DeleteCommand.Connection.Close();

        myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM Время", myOleDbConnection);
        myDataAdapter.SelectCommand.Connection.Open();
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.SelectCommand.Connection.Close();

        myDataSet.Tables["Время"].Clear();
        myDataAdapter.Fill(myDataSet, "Время");
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
        obj_connect = null;
      }
    }

    
  }
}
