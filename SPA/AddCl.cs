using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Data.OleDb;

namespace SPA
{
  public partial class AddCl : Form
  {
    OleDbConnection myOleDbConnection;
    OleDbDataAdapter myDataAdapter;
    DataSet myDataSet;
    public OleDbConnection obj_connect = null;
    string connectionString;

    public AddCl()
    {
      InitializeComponent();
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

      myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Склад", myOleDbConnection);
      myDataAdapter.SelectCommand.ExecuteNonQuery();
      myDataAdapter.Fill(myDataSet, "Склад");

      myDataAdapter.SelectCommand.Connection.Close();
    }

    private void AddCl_Load(object sender, EventArgs e)
    {
      ToolTip t = new ToolTip();
      t.SetToolTip(this.button1, "Добавить клиента");     
      t.SetToolTip(this.button2, "Выход");
    }

    private void button1_Click(object sender, EventArgs e)
    {
      string cmd = "INSERT INTO Клиенты  VALUES (" + textBox1.Text + ",'" + textBox2.Text + "','" + textBox3.Text + "', '" + textBox4.Text + "','" + maskedTextBox1.Text + "','" + textBox5.Text + "','" + textBox6.Text + "', '" + textBox7.Text + "', '" + textBox7.Text + "' )";
      try
      {
        myDataAdapter.InsertCommand = new OleDbCommand(cmd, myOleDbConnection);

        myDataAdapter.InsertCommand.Connection.Open();
        myDataAdapter.InsertCommand.ExecuteNonQuery();
        //MessageBox.Show(myDataAdapter.InsertCommand.CommandText);
        myDataAdapter.InsertCommand.Connection.Close();

        myDataAdapter.SelectCommand = new OleDbCommand("Select * FROM Клиенты", myOleDbConnection);
        myDataAdapter.SelectCommand.Connection.Open();
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.SelectCommand.Connection.Close();

        myDataSet.Tables["Клиенты"].Clear();
        myDataAdapter.Fill(myDataSet, "Клиенты");
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
        obj_connect = null;
      }
    }

    private void button2_Click(object sender, EventArgs e)
    {
      this.Close();
    }

    private void button1_Move(object sender, EventArgs e)
    {
  
    }
  }
}
