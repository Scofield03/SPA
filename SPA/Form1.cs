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
  public partial class Form1 : Form
  {
    OleDbConnection myOleDbConnection;
    OleDbDataAdapter myDataAdapter;
    DataSet myDataSet;
    System.Timers.Timer timer;

    AddLenTime AddLenTime;
    AddTime AddTime;
    Login Login;
    BossAdd BossAdd;
    AddCl AddCl;

    public OleDbConnection obj_connect = null;
    string connectionString;
    public Form1()
    {
      InitializeComponent();
    }

    private void Form1_Load(object sender, EventArgs e)
    {
      ToolTip t = new ToolTip();
      t.SetToolTip(this.button3, "Вход администратора");//SetToolTipTitle(this.button3, "Выход");
      t.SetToolTip(this.button11, "Обновление");
      t.SetToolTip(this.button10, "Редактировать время приема");
      t.SetToolTip(this.button1, "Поиск клиентов для выбранного специалиста");
      t.SetToolTip(this.button18, "Минимизация окна");
      t.SetToolTip(this.label3, "Время");
      t.SetToolTip(this.button8, "Обновление");
      t.SetToolTip(this.button2, "Добавить клиента в базу");
      t.SetToolTip(this.button6, "Записать клиента на прием");
      t.SetToolTip(this.button9, "Удалить клиента");
      t.SetToolTip(this.button15, "Минимизация окна");
      t.SetToolTip(this.button16, "Открыть навигацию");
      t.SetToolTip(this.button17, "Закрыть навигацию");
      t.SetToolTip(this.button14, "Найти клиента по полису");
      t.SetToolTip(this.button4, "Найти клиента по фамилии");
      t.SetToolTip(this.button5, "Показать результат навигации");
      t.SetToolTip(this.button7, "Минимизация окна");

      

      System.Windows.Forms.Timer T = new System.Windows.Forms.Timer();
      T.Interval = 1000; //Выполнять каждые 10 секунд
      T.Tick += new EventHandler(T_Tick);
      T.Enabled = true;

      this.Width = 1094;
      this.Height = 597;
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


      myDataAdapter.SelectCommand.Connection.Close();

      this.dataGridView1.DataSource = myDataSet.Tables[0].DefaultView;
      this.dataGridView2.DataSource = myDataSet.Tables[1].DefaultView;
     // this.dataGridView3.DataSource = myDataSet.Tables["Расписание"].DefaultView;
    // this.dataGridView4.DataSource = myDataSet.Tables["Расписание"].DefaultView;

      this.dataGridView1.Columns["ID_Персонала"].Visible = false;
     // this.dataGridView3.Columns["ID_расписания"].Visible = false;

     // this.dataGridView4.Columns["ID_расписания"].Visible = false;
     // this.dataGridView4.Columns["ID_расписания"].Visible = false;
      //this.dataGridView4.Columns["ID_ингредиента"].Visible = false;

      comboBox1.DataSource = myDataSet.Tables["Персонал"].DefaultView;
      comboBox1.DisplayMember = "Фамилия";

      comboBox2.DataSource = myDataSet.Tables["Процедуры"].DefaultView;
      comboBox2.DisplayMember = "Название";

      comboBox3.DataSource = myDataSet.Tables["Клиенты"].DefaultView;
      comboBox3.DisplayMember = "Фамилия";

      comboBox4.DataSource = myDataSet.Tables["Клиенты"].DefaultView;
      comboBox4.DisplayMember = "Полис";


   

      this.dataGridView1.DataSource = myDataSet.Tables[0].DefaultView;
      
    }

    private void textBox8_TextChanged(object sender, EventArgs e)
    {

    }

    private void button9_Click(object sender, EventArgs e)
    {
      try
      {
        myDataAdapter.DeleteCommand = new OleDbCommand("DELETE FROM Клиенты WHERE Полис=" + dataGridView2.SelectedRows[0].Cells[0].Value, myOleDbConnection);

        myDataAdapter.DeleteCommand.Connection.Open();
        myDataAdapter.DeleteCommand.ExecuteNonQuery();
        MessageBox.Show(myDataAdapter.DeleteCommand.CommandText);
        myDataAdapter.DeleteCommand.Connection.Close();

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

    private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
    {

    }

    private void button14_Click(object sender, EventArgs e)
    {
      try
      {
        //textBox6.this.dataGridView2.DataSource = myDataSet.Tables["Клиенты"].DefaultView;
        myDataSet.Tables["Клиенты"].Clear();

        myDataSet.Tables["Клиенты"].Clear();
        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * From Клиенты WHERE Полис=" + textBox9.Text, myOleDbConnection);
        myDataAdapter.SelectCommand.Connection.Open();
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Клиенты");
        myDataAdapter.SelectCommand.Connection.Close();
        textBox9.Clear();
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
        obj_connect = null;
      }
    }

    private void button16_Click(object sender, EventArgs e)
    {
      groupBox1.Visible = true;
      this.Width = 1258;
      this.Height = 597;
 

    }  

    private void button17_Click(object sender, EventArgs e)
    {
      groupBox1.Visible = false;
      this.Width = 1094;
      this.Height = 597;
    }

    private void tabPage2_Click(object sender, EventArgs e)
    {

    }

    private void button11_Click(object sender, EventArgs e)
    {
      try
      {
        this.dataGridView1.DataSource = myDataSet.Tables[0].DefaultView;
        this.dataGridView1.Columns["ID_Персонала"].Visible = false;
        myDataSet.Tables["Персонал"].Clear();
        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * From Персонал ", myOleDbConnection);
        myDataAdapter.SelectCommand.Connection.Open();
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Персонал");
        myDataAdapter.SelectCommand.Connection.Close();
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
        obj_connect = null;
      }
    }

    private void button10_Click(object sender, EventArgs e)
    {

        //this.WindowState = FormWindowState.Minimized;

        //this.Hide();
        AddLenTime = new AddLenTime();
        AddLenTime.Owner = this;
        AddLenTime.ShowDialog();
      
    }

    private void timer1_Tick(object sender, EventArgs e)
    {
      label3.Text = DateTime.Now.ToShortTimeString();

    }

    private void label3_Click(object sender, EventArgs e)
    {

    }

    private void button6_Click(object sender, EventArgs e)
    {
        AddTime = new AddTime();
        AddTime.Owner = this;
        AddTime.ShowDialog();
    }

    private void button1_Click(object sender, EventArgs e)
    {
        try
        {
            this.dataGridView1.DataSource = myDataSet.Tables["Расписание"].DefaultView;

            this.dataGridView1.Columns["ID_расписания"].Visible = false;

            myDataSet.Tables["Расписание"].Clear();
            myDataAdapter.SelectCommand = new OleDbCommand("SELECT * From Расписание Where Специалист='" + textBox8.Text + "'", myOleDbConnection);
            myDataAdapter.SelectCommand.Connection.Open();
            myDataAdapter.SelectCommand.ExecuteNonQuery();
            myDataAdapter.Fill(myDataSet, "Расписание");
            myDataAdapter.SelectCommand.Connection.Close();

        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
            obj_connect = null;
            //AddCl.myDataAdapter.SelectCommand.Connection.Open();
        }
    }

    private void button3_Click(object sender, EventArgs e)
    {
        Login = new Login();
        Login.Owner = this;
        Login.ShowDialog();
        
        
        
    }

    private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
    {

    }

    private void button18_Click(object sender, EventArgs e)
    {
      this.WindowState = FormWindowState.Minimized;
       
    }

    private void button2_Click(object sender, EventArgs e)
    {
      AddCl = new AddCl();
      AddCl.Owner = this;
      
      AddCl.ShowDialog(); 
    }

    private void button15_Click(object sender, EventArgs e)
    {
     
    }

    private void button4_Click(object sender, EventArgs e)
    {
      try
      {
        myDataSet.Tables["Клиенты"].Clear();
        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * From Клиенты WHERE Фамилия='" + textBox2.Text + "'", myOleDbConnection);
        myDataAdapter.SelectCommand.Connection.Open();
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Клиенты");
        myDataAdapter.SelectCommand.Connection.Close();
        textBox2.Clear();
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
        obj_connect = null;
      }
        
    }

    private void button12_Click(object sender, EventArgs e)
    {
      
    }

    private void button19_Click(object sender, EventArgs e)
    {
   
    }   

    private void T_Tick(object sender, EventArgs e)
          {
            sql();
          }
 
    private void sql()
    {
      this.dataGridView4.DataSource = myDataSet.Tables["Расписание"].DefaultView;

      this.dataGridView4.Columns["ID_расписания"].Visible = false;

      myDataSet.Tables["Расписание"].Clear();
      myDataAdapter.SelectCommand = new OleDbCommand("SELECT * From Расписание Where (С='" + label3.Text + "' AND Дата= '"+DateTime.Now.ToShortDateString()+"')", myOleDbConnection);
      myDataAdapter.SelectCommand.Connection.Open();
      myDataAdapter.SelectCommand.ExecuteNonQuery();
      myDataAdapter.Fill(myDataSet, "Расписание");
      myDataAdapter.SelectCommand.Connection.Close();
    }

    private void button20_Click(object sender, EventArgs e)
    {

    }

    private void button3_MouseMove(object sender, MouseEventArgs e)
    {
    

    }

    private void button3_MouseLeave(object sender, EventArgs e)
    {
     
    }

    private void button8_Click(object sender, EventArgs e)
    {
      try
      {

        myDataSet.Tables["Клиенты"].Clear();
        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * From Клиенты ", myOleDbConnection);
        myDataAdapter.SelectCommand.Connection.Open();
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Клиенты");
        myDataAdapter.SelectCommand.Connection.Close();
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
        obj_connect = null;
      }
    }

    private void label2_Click(object sender, EventArgs e)
    {

    }

    private void button15_Click_1(object sender, EventArgs e)
    {
      this.WindowState = FormWindowState.Minimized;
    }

    private void trackBar1_Scroll(object sender, EventArgs e)
    {
      Form1.ActiveForm.Opacity = trackBar1.Value / (double)trackBar1.Maximum;
    }

    private void label4_Click(object sender, EventArgs e)
    {

    }

    private void dateTimePicker1_ValueChanged_1(object sender, EventArgs e)
    {

    }

    private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
    {

    }

    private void button5_Click(object sender, EventArgs e)
    {
      try
      {
        this.dataGridView3.DataSource = myDataSet.Tables["Расписание"].DefaultView;
        this.dataGridView3.Columns["ID_расписания"].Visible = false;
        myDataSet.Tables["Расписание"].Clear();
        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * From Расписание ", myOleDbConnection);
        myDataAdapter.SelectCommand.Connection.Open();
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Расписание");
        myDataAdapter.SelectCommand.Connection.Close();
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
        obj_connect = null;
      }
    }

    private void button7_Click(object sender, EventArgs e)
    {
      this.WindowState = FormWindowState.Minimized;
    }

    private void button12_Click_1(object sender, EventArgs e)
    {
        MessageBox.Show(DateTime.Now.ToShortDateString(), "");
    }
  }
}
