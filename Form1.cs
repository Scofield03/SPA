﻿using System;
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
    System.Windows.Forms.Timer T = new System.Windows.Forms.Timer();

    public OleDbConnection obj_connect = null;
    string connectionString;
    double qw = 60;

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
      t.SetToolTip(this.button15, "Навигация по персаналу");
      t.SetToolTip(this.label3, "Время");
      t.SetToolTip(this.button8, "Обновление");
      t.SetToolTip(this.button13, "Добавить клиента в базу");
     // t.SetToolTip(this.button6, "Записать клиента на прием");
      t.SetToolTip(this.button9, "Удалить клиента");
      t.SetToolTip(this.button19, "Записать на прием");
      t.SetToolTip(this.button20, "Вернуться назад");
      t.SetToolTip(this.button15, "Навигация по персоналу");
    
      //t.SetToolTip(this.button16, "Открыть навигацию");
      //t.SetToolTip(this.button17, "Закрыть навигацию");
      //t.SetToolTip(this.button14, "Найти клиента по полису");
      t.SetToolTip(this.button4, "Найти клиента по фамилии");
      t.SetToolTip(this.button5, "Показать результат навигации");
      t.SetToolTip(this.button7, "Минимизация окна");

      

      //System.Windows.Forms.Timer T = new System.Windows.Forms.Timer();
      T.Interval = 1000; //Выполнять каждые 10 секунд
      T.Tick += new EventHandler(T_Tick);
      T.Enabled = true;   
      
      //this.Width = 1094;
      //this.Height = 597;1125; 937
      this.Width = 1125;
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

      myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM spa_процедуры", myOleDbConnection);
      myDataAdapter.SelectCommand.ExecuteNonQuery();
      myDataAdapter.Fill(myDataSet, "spa_процедуры");

      myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание", myOleDbConnection);
      myDataAdapter.SelectCommand.ExecuteNonQuery();
      myDataAdapter.Fill(myDataSet, "Расписание");

      myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание", myOleDbConnection);
      myDataAdapter.SelectCommand.ExecuteNonQuery();
      myDataAdapter.Fill(myDataSet, "Расписание2");

      myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание", myOleDbConnection);
      myDataAdapter.SelectCommand.ExecuteNonQuery();
      myDataAdapter.Fill(myDataSet, "Расписание12");

      myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание", myOleDbConnection);
      myDataAdapter.SelectCommand.ExecuteNonQuery();
      myDataAdapter.Fill(myDataSet, "Расписание1");

      myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание", myOleDbConnection);
      myDataAdapter.SelectCommand.ExecuteNonQuery();
      myDataAdapter.Fill(myDataSet, "Расписание3");

      myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание", myOleDbConnection);
      myDataAdapter.SelectCommand.ExecuteNonQuery();
      myDataAdapter.Fill(myDataSet, "Расписание31");

      myDataAdapter.SelectCommand = new OleDbCommand("SELECT Расписание.Дата, Клиенты.Фамилия, Процедуры.Название, Персонал.Фамилия, Расписание.С, Расписание.По FROM Персонал, Процедуры INNER JOIN (Клиенты INNER JOIN Расписание ON Клиенты.Полис = Расписание.Клиент) ON Процедуры.ID_Процедуры = Расписание.Процедура WHERE (((Персонал.ID_Персонала)=[Расписание].[Специалист]));", myOleDbConnection);
      myDataAdapter.SelectCommand.ExecuteNonQuery();
      myDataAdapter.Fill(myDataSet, "Расписание41");




      myDataAdapter.SelectCommand.Connection.Close();

      this.dataGridView1.DataSource = myDataSet.Tables[0].DefaultView;
      this.dataGridView2.DataSource = myDataSet.Tables[1].DefaultView;
      this.dataGridView3.DataSource = myDataSet.Tables["Расписание31"].DefaultView;

      this.dataGridView5.DataSource = myDataSet.Tables["Расписание1"].DefaultView;
      this.dataGridView4.DataSource = myDataSet.Tables["Расписание"].DefaultView;

      this.dataGridView1.Columns["ID_Персонала"].Visible = false;
      this.dataGridView3.Columns["ID_расписания"].Visible = false;

     this.dataGridView4.Columns["ID_расписания"].Visible = false;
      this.dataGridView5.Columns["ID_расписания"].Visible = false;
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
      //Form1.ActiveForm.Opacity = 0.5;
    }

    //private void T_Tick1(object Sender, EventArgs e)
    //{
     
    //  while (qw != 0)
    //  {
    //    qw = qw - 1;
    //   // cook();
    //  }
      
      
   // } 

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
      T.Enabled = false;
        //this.Location = new Point(0, 0);
        this.Width = 1125;
        this.Height = 597;


        dataGridView1.Location = new Point(437, 0);
        dataGridView1.Width = 635;
        dataGridView1.Height = 240
           ;

        //dataGridView6.Location = new Point (3, 420);
      //  groupBox3.Location = new Point(905, 507);

       // dataGridView6.Width= 894;
        //dataGridView6.Height = 413;

        button3.Visible = false;
        button10.Visible = false;
        button15.Visible = false;
        button19.Visible = true;
        button20.Visible = true;
        dataGridView6.Visible = true;
        dateTimePicker2.Visible = true;

        radioButton1.Visible = true;
        radioButton2.Visible = true;
        radioButton3.Visible = true;

        groupBox2.Enabled = false;
        groupBox2.Visible = true;
        groupBox3.Visible = true;

        comboBox6.Visible = true;
        comboBox5.Visible = true;
        comboBox5.Enabled = false;
        comboBox6.Enabled = false;
        comboBox7.Enabled = false;
        comboBox7.Visible = true;
        comboBox9.Visible = true;
        comboBox10.Visible = true;
        comboBox11.Visible = true;
        comboBox12.Visible = true;
        label24.Visible = true;
        label25.Visible = true;
        label26.Visible = true;
        label27.Visible = true;
        label28.Visible = true;
        

        label14.Visible = true;
        label13.Visible = true;
        label12.Visible = true;
        label11.Visible = true;
        string cmd;
     // if(comboBox6.SelectedIndex=== -1 || comboBox6.Text == string.Empty & comboBox7.SelectedIndex==-1|| comboBox7.Text == string.Empty & comboBox5.SelectedIndex==-1 || comboBox5.Text == string.Empty) 
       //cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Процедура = '" + comboBox2.Text + "' AND Специалист = '" + comboBox1.Text + "' AND Полис = " + comboBox4.Text + ")";
        myOleDbConnection = new OleDbConnection(connectionString);


       // myOleDbConnection = new OleDbConnection(connectionString);
        // myDataAdapter = new System.Data.OleDb.OleDbDataAdapter("SELECT * FROM Время WHERE (flag=True)", myOleDbConnection);
        myDataAdapter = new System.Data.OleDb.OleDbDataAdapter("SELECT * FROM Время", myOleDbConnection);
        myDataSet = new DataSet("Время12");
        myDataAdapter.Fill(myDataSet, "Время12");
        myDataAdapter.SelectCommand.Connection.Close();

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Клиенты", myOleDbConnection);
        myDataAdapter.SelectCommand.Connection.Open();
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Клиенты");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Персонал", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Персонал");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Персонал WHERE (Not Специальность='СПА-мастер')", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "NonSPA_Персонал");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Персонал WHERE (Специальность = 'СПА-мастер')", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "SPA_Персонал");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Процедуры", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Процедуры");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Специальности", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Специальности");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM spa_процедуры", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "spa_процедуры");


        myDataAdapter.SelectCommand.Connection.Close();

        this.dataGridView6.DataSource = myDataSet.Tables["Клиенты"].DefaultView;
        this.dataGridView1.DataSource = myDataSet.Tables["Время12"].DefaultView;

        //this.dataGridView1.Columns["ID_расписания"].Visible = false;
        //this.dataGridView1.Columns["Клиент"].Visible = false;
        //this.dataGridView1.Columns["Процедура"].Visible = false;
        //this.dataGridView1.Columns["Специалист"].Visible = false;

        this.dataGridView6.Columns["Телефон"].Visible = false;
        this.dataGridView6.Columns["Город"].Visible = false;
        this.dataGridView6.Columns["Дом"].Visible = false;
        this.dataGridView6.Columns["Улица"].Visible = false;
        this.dataGridView6.Columns["Квартира"].Visible = false;

        //comboBox5.DataSource = myDataSet.Tables["Процедуры"].DefaultView;
       // comboBox5.DisplayMember = "Название";

        comboBox6.DataSource = myDataSet.Tables["Персонал"].DefaultView;
        comboBox6.DisplayMember = "Фамилия";

        comboBox8.DataSource = myDataSet.Tables["Клиенты"].DefaultView;
        comboBox8.DisplayMember = "Полис";

        //comboBox7.DataSource = myDataSet.Tables["spa_процедуры"].DefaultView;
        //comboBox7.DisplayMember = "Название";

        //string cmd = "SELECT * FROM Время WHERE ((Процедура= '" + comboBox1.Text + "') AND ([_Дата]='" + dateTimePicker1.Value.ToShortDateString() + "') AND (flag=True))";
        //myDataAdapter.SelectCommand = new OleDbCommand(cmd, myOleDbConnection);
        //myDataAdapter.SelectCommand.Connection.Open();
        //myDataAdapter.SelectCommand.ExecuteNonQuery();
        //myDataAdapter.SelectCommand.Connection.Close();

        //myDataSet.Tables["Время"].Clear();        
        //myDataAdapter.Fill(myDataSet, "Время"); 
        //this.dataGridView1.DataSource = myDataSet.Tables["Время"].DefaultView;
    }

    private void timer1_Tick(object sender, EventArgs e)
    {
     label3.Text = DateTime.Now.ToShortTimeString();

    }

    private void label3_Click(object sender, EventArgs e)
    {

    }



    private void button1_Click(object sender, EventArgs e)
    {
        try
        {
            int a = 0;
            string cc = "SELECT Персонал.ID_Персонала FROM Персонал WHERE (((Персонал.Фамилия)='" + textBox8.Text + "'));";
            OleDbConnection myConn = new OleDbConnection(connectionString);
            myConn.Open();
            this.dataGridView1.DataSource = myDataSet.Tables["Расписание41"].DefaultView;
            OleDbCommand cmd = new OleDbCommand(cc, myConn);
            a = Convert.ToInt32(cmd.ExecuteScalar());
            myDataSet.Tables["Расписание41"].Clear();
            myDataAdapter.SelectCommand = new OleDbCommand("SELECT Расписание.Дата, Клиенты.Фамилия, Процедуры.Название, Персонал.Фамилия, Расписание.С, Расписание.По FROM Персонал, Процедуры INNER JOIN (Клиенты INNER JOIN Расписание ON Клиенты.Полис = Расписание.Клиент) ON Процедуры.ID_Процедуры = Расписание.Процедура WHERE (((Персонал.ID_Персонала)=[Расписание].[Специалист] AND Специалист=" + a + "));", myOleDbConnection);
            myDataAdapter.SelectCommand.Connection.Open();
            myDataAdapter.SelectCommand.ExecuteNonQuery();
            myDataAdapter.Fill(myDataSet, "Расписание41");
            myDataAdapter.SelectCommand.Connection.Close();

        }
        catch (Exception ex)
        {

            MessageBox.Show(ex.Message);
            obj_connect = null;
            //AddCl.myDataAdapter.SelectCommand.Connection.Open();
        }
        textBox8.Text = "Введите фамилию";
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
           label3.Text = DateTime.Now.ToShortTimeString();
            sql();
          }
 
    private void sql()
    {

      

     this.dataGridView4.Columns["ID_расписания"].Visible = false;
     this.dataGridView5.Columns["ID_расписания"].Visible = false;

     myDataSet.Tables["Расписание"].Clear();
     myDataAdapter.SelectCommand.Connection.Close();
     myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание WHERE (С ='" + label3.Text + "' AND Дата= '" + DateTime.Now.ToShortDateString() + "')", myOleDbConnection);
     myDataAdapter.SelectCommand.Connection.Open();
     myDataAdapter.SelectCommand.ExecuteNonQuery();
     myDataAdapter.Fill(myDataSet, "Расписание");
     myDataAdapter.SelectCommand.Connection.Close();
     this.dataGridView4.DataSource = myDataSet.Tables["Расписание"].DefaultView;


      //this.dataGridView5.Columns["ID_расписания"].Visible = false;

      myDataSet.Tables["Расписание1"].Clear();
      myDataAdapter.SelectCommand.Connection.Close();
      myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание WHERE (По ='" + label3.Text + "' AND Дата= '" + DateTime.Now.ToShortDateString() + "')", myOleDbConnection);
      myDataAdapter.SelectCommand.Connection.Open();
      myDataAdapter.SelectCommand.ExecuteNonQuery();
      myDataAdapter.Fill(myDataSet, "Расписание1");
      myDataAdapter.SelectCommand.Connection.Close();
      this.dataGridView5.DataSource = myDataSet.Tables["Расписание1"].DefaultView;
   
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
        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Клиенты ", myOleDbConnection);
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
        string cmd = "", pers = " ", proc = " ", fio = " ";

        OleDbConnection myConn = new OleDbConnection(connectionString);
        myConn.Open();
        if (comboBox1.SelectedIndex != -1 || comboBox1.Text != string.Empty)
        {
            cmd = "SELECT ID_Персонала FROM Персонал WHERE (Фамилия = '" + comboBox1.Text + "')";
            OleDbCommand cmd1 = new OleDbCommand(cmd, myConn);

            pers = cmd1.ExecuteScalar().ToString();
            //MessageBox.Show(pers);
        }

        if (comboBox2.SelectedIndex != -1 || comboBox2.Text != string.Empty)
        {
            cmd = "SELECT ID_Процедуры FROM Процедуры WHERE (Название = '" + comboBox2.Text + "')";
            OleDbCommand cmd1 = new OleDbCommand(cmd, myConn);
            proc = cmd1.ExecuteScalar().ToString();
            //MessageBox.Show(proc);
        }

        if (comboBox3.SelectedIndex != -1 || comboBox3.Text != string.Empty)
        {
            cmd = "SELECT Полис FROM Клиенты WHERE (Фамилия = '" + comboBox3.Text + "')";
            OleDbCommand cmd1 = new OleDbCommand(cmd, myConn);
            fio = cmd1.ExecuteScalar().ToString();
           // MessageBox.Show(fio);
        }
        myConn.Close();
        //try
        // {
        this.dataGridView3.DataSource = myDataSet.Tables["Расписание31"].DefaultView;
        // this.dataGridView3.Columns["ID_расписания"].Visible = false;
        myDataSet.Tables["Расписание31"].Clear();
        if (comboBox1.SelectedIndex != -1 || comboBox1.Text != string.Empty)
        {
            if (comboBox2.SelectedIndex != -1 || comboBox2.Text != string.Empty)
            {
                if (comboBox3.SelectedIndex != -1 || comboBox3.Text != string.Empty)
                {
                    if (comboBox4.SelectedIndex != -1 || comboBox4.Text != string.Empty)
                        //MessageBox.Show("1");
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Процедура = " + proc + " AND Специалист = " + pers + " AND Полис = " + comboBox4.Text + " AND Клиент = " + fio + ")";
                    else
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Процедура = " + proc + " AND Специалист = " + pers + " AND Клиент = " + fio + ")";

                }
                else
                {
                    if (comboBox4.SelectedIndex != -1 && comboBox4.Text != string.Empty)
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Процедура = " + proc + " AND Специалист = " + pers + " AND Полис = " + comboBox4.Text + ")";
                    else
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Процедура = " + proc + " AND Специалист = " + pers + ")";
                }
            }
            else
            {
                if (comboBox3.SelectedIndex != -1 || comboBox3.Text != string.Empty)
                {
                    if (comboBox4.SelectedIndex != -1 || comboBox4.Text != string.Empty)
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Специалист = " + pers + " AND Полис = " + comboBox4.Text + " AND Клиент = " + fio + ")";
                    else
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Специалист = " + pers + " AND Клиент = " + fio + ")";

                }
                else
                {
                    if (comboBox4.SelectedIndex != -1 && comboBox4.Text != string.Empty)
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Специалист = " + pers + " AND Полис = " + comboBox4.Text + ")";
                    else
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Специалист = " + pers + ")";
                }
            }
        }
        else
        {
            if (comboBox2.SelectedIndex != -1 || comboBox2.Text != string.Empty)
            {
                if (comboBox3.SelectedIndex != -1 || comboBox3.Text != string.Empty)
                {
                    if (comboBox4.SelectedIndex != -1 || comboBox4.Text != string.Empty)
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Процедура = " + proc + " AND Полис = " + comboBox4.Text + " AND Клиент = " + fio + ")";
                    else
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Процедура = " + proc + " AND Клиент = " + fio + ")";

                }
                else
                {
                    if (comboBox4.SelectedIndex != -1 && comboBox4.Text != string.Empty)
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Процедура = " + proc + " AND Полис = " + comboBox4.Text + ")";
                    else
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Процедура = " + proc + ")";
                }
            }
            else
            {
                if (comboBox3.SelectedIndex != -1 || comboBox3.Text != string.Empty)
                {
                    if (comboBox4.SelectedIndex != -1 || comboBox4.Text != string.Empty)
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Полис = " + comboBox4.Text + " AND Клиент = " + fio + ")";
                    else
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "' AND Клиент = " + fio + ")";
                }
                else
                {
                    if (comboBox4.SelectedIndex != -1 && comboBox4.Text != string.Empty)
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "'AND Полис = " + comboBox4.Text + ")";
                    else
                        cmd = "SELECT Дата, Процедура, Специалист, Клиент, Полис FROM Расписание, Клиенты WHERE (Дата = '" + dateTimePicker1.Value.ToShortDateString() + "')";
                }
            }
        }

        myDataAdapter.SelectCommand = new OleDbCommand(cmd, myOleDbConnection);
        myDataAdapter.SelectCommand.Connection.Open();
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Расписание31");
        myDataAdapter.SelectCommand.Connection.Close();
        // }
        /* catch (Exception ex)
         {
             MessageBox.Show(ex.Message);
             obj_connect = null;
         }*/
    }

    private void button7_Click(object sender, EventArgs e)
    {
      this.WindowState = FormWindowState.Minimized;
    }

    private void button12_Click_1(object sender, EventArgs e)
    {
        MessageBox.Show(DateTime.Now.ToShortDateString(), "");
    }

    private void textBox8_MouseClick(object sender, MouseEventArgs e)
    {
        textBox8.Text = "";
    }

    private void label11_Click(object sender, EventArgs e)
    {

    }

   

    private void radioButton1_CheckedChanged(object sender, EventArgs e)
    {
        if (radioButton1.Checked == true)
        {
            comboBox6.DataSource = myDataSet.Tables["NonSPA_Персонал"].DefaultView;
            comboBox6.DisplayMember = "Фамилия";
        }
       /* else
        {
            comboBox6.DataSource = myDataSet.Tables["SPA_Персонал"].DefaultView; 
            comboBox6.DisplayMember = "Фамилия";
        }*/
        comboBox5.Enabled = true;
        comboBox6.Enabled = true;
        comboBox7.Enabled = false;
        
    }

    private void radioButton2_CheckedChanged(object sender, EventArgs e)
    {
        if (radioButton2.Checked == true)
        {
            comboBox6.DataSource = myDataSet.Tables["SPA_Персонал"].DefaultView;
            comboBox6.DisplayMember = "Фамилия";
            
        }
        /* else
         {
 comboBox6.DataSource = myDataSet.Tables["NonSPA_Персонал"].DefaultView;
             comboBox6.DisplayMember = "Фамилия";
         }*/
        comboBox7.Enabled = true;
        comboBox6.Enabled = true;
        comboBox5.Enabled = false;
    }

    private void radioButton3_CheckedChanged(object sender, EventArgs e)
    {
        groupBox2.Enabled = true;
        groupBox2.Visible = true;
       // dataGridView6.Location = new Point (3,538);
      //  groupBox3.Location = new Point(905, 507);
       // dataGridView6.Width = 894;
       // dataGridView6.Height = 300;
        radioButton3.Checked = false;
    }

    private void button13_Click(object sender, EventArgs e)
    {
        groupBox2.Enabled   = false;
      //  dataGridView6.Location = new Point(3, 420);
      //  groupBox3.Location = new Point(905, 507);
      //  dataGridView6.Width = 894;
       // dataGridView6.Height = 413;


        string cmd = "INSERT INTO Клиенты  VALUES (" + textBox11.Text + ",'" + textBox10.Text + "','" + textBox3.Text + "', '" + textBox4.Text + "','" + maskedTextBox1.Text + "','" + textBox5.Text + "','" + textBox6.Text + "', '" + textBox7.Text + "', '" + textBox1.Text + "' )";
        try
        {
            myDataAdapter.InsertCommand = new OleDbCommand(cmd, myOleDbConnection);

            myDataAdapter.InsertCommand.Connection.Open();
            myDataAdapter.InsertCommand.ExecuteNonQuery();
            //MessageBox.Show(myDataAdapter.InsertCommand.CommandText);
            myDataAdapter.InsertCommand.Connection.Close();           

             

           // cmd ="SELECT * FROM Клиенты WHERE Полис= " + textBox11.Text + " ";
            try
            {
              myDataSet.Tables["Клиенты"].Clear();
              myDataSet.Tables["Клиенты"].Clear();
              myDataAdapter.SelectCommand = new OleDbCommand("SELECT * From Клиенты WHERE Полис=" + textBox11.Text, myOleDbConnection);
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
            MessageBox.Show("Клиент добавлен в базу", "Внимание!");

            textBox11.Clear();
            textBox10.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            textBox1.Clear();
            maskedTextBox1.Clear();

        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
            obj_connect = null;
        }
        
    }

    private void button19_Click_1(object sender, EventArgs e)
    {
     int ID_PROC=0;
     int ID_PERS=0;
        try
        {

        //   // if (dataGridView1.SelectedRows.Count != 1)
        //   //     MessageBox.Show("Выберите свободное время в таблице!", "ОШИБКА!", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //   // else
        ////    {



        // //SqlConnection obj_connect1 = new SqlConnection("Data Source=(spa.mdb)");
        try
         {

          string comand = "SELECT  ID_Процедуры FROM Процедуры WHERE (((Название='" + comboBox5.Text + "')))";
          obj_connect = new OleDbConnection(connectionString);
          OleDbCommand com = new OleDbCommand(comand, obj_connect);
          obj_connect.Open();
         // com.ExecuteNonQuery();
          ID_PROC = (int)com.ExecuteScalar(); //reader.GetInt32(0); 
       //   obj_connect.Close();


          comand = "SELECT  ID_Персонала FROM Персонал WHERE (((Фамилия='" + comboBox6.Text + "')))";
         // obj_connect = new OleDbConnection(connectionString);
          OleDbCommand com1 = new OleDbCommand(comand, obj_connect);
          ///com.ExecuteNonQuery();
          ID_PERS = (int)com1.ExecuteScalar();
          obj_connect.Close();
         }
       catch (Exception ex)
         {
          MessageBox.Show(ex.Message);
          obj_connect = null;
         }
        

               
         string cmd = "INSERT INTO Расписание (Клиент, Процедура, Специалист,С,По,Дата)  VALUES (" + Convert.ToInt32(comboBox8.Text) + "," + ID_PROC + "," + ID_PERS + ",'" + comboBox12.SelectedItem.ToString() + ":" + comboBox11.SelectedItem.ToString() + "','" + comboBox10.SelectedItem.ToString() + ":" + comboBox9.SelectedItem.ToString() + "','"+dateTimePicker2.Value.ToShortDateString()+"')";
        // myDataAdapter.InsertCommand.Connection.Close();
         myDataAdapter.InsertCommand = new OleDbCommand(cmd, myOleDbConnection);                
                myDataAdapter.InsertCommand.Connection.Open();
                MessageBox.Show(myDataAdapter.InsertCommand.CommandText);
                myDataAdapter.InsertCommand.ExecuteNonQuery();                
                myDataAdapter.InsertCommand.Connection.Close();

                myDataSet.Tables["Расписание12"].Clear();
                myDataSet.Tables["Расписание12"].Clear();
                myDataAdapter.SelectCommand = new OleDbCommand("SELECT * From Расписание ", myOleDbConnection);
                myDataAdapter.SelectCommand.Connection.Open();
                myDataAdapter.SelectCommand.ExecuteNonQuery();
                myDataAdapter.Fill(myDataSet, "Расписание12");
                myDataAdapter.SelectCommand.Connection.Close();
                //cmd = "UPDATE Время SET flag = True WHERE (Дата='" + dataGridView1.SelectedRows[0].Cells[5].Value.ToString() + "')";
                //myDataAdapter.UpdateCommand = new OleDbCommand(cmd, myOleDbConnection);
                //myDataAdapter.UpdateCommand.Connection.Open();
                //myDataAdapter.UpdateCommand.ExecuteNonQuery();
                //MessageBox.Show(myDataAdapter.UpdateCommand.CommandText);
                //myDataAdapter.UpdateCommand.Connection.Close();


            }
      // }
        catch (Exception ex)
        {

            MessageBox.Show(ex.Message);

        }
    }

    private void button20_Click_1(object sender, EventArgs e)
    {

      T.Enabled = true;
      this.Location = new Point(100, 100);
        this.Width = 1125;
        this.Height = 597;
        dataGridView1.Width = 894;
        dataGridView1.Height = 499;
        dataGridView1.Location = new Point(6,6);
        dataGridView6.Visible = false;
      //  groupBox3.Location = new Point(905, 507);

      //  dataGridView6.Width = 894;
     //   dataGridView6.Height = 413;

        button3.Visible = true;
        button10.Visible = true;
        button15.Visible = true;
        button19.Visible = false;
        button20.Visible = false;
        groupBox2.Visible = false;
        dateTimePicker2.Visible = false;

        radioButton1.Visible = false;
        radioButton2.Visible = false;
        radioButton3.Visible = false;

        groupBox2.Enabled = false;
        groupBox2.Visible = false;

        groupBox3.Visible = false;

        comboBox6.Visible = false;
        comboBox5.Visible = false;
        comboBox5.Enabled = false;
        comboBox7.Enabled = false;
        comboBox7.Visible = false;
        comboBox9.Visible = false;
        comboBox10.Visible = false;
        comboBox11.Visible = false;
        comboBox12.Visible = false;

        label24.Visible = false;
        label25.Visible = false;
        label26.Visible = false;
        label27.Visible = false;
        label28.Visible = false;


        label14.Visible = false;
        label13.Visible = false;
        label12.Visible = false;
        label11.Visible = false;

        myOleDbConnection = new OleDbConnection(connectionString);
        myDataAdapter = new System.Data.OleDb.OleDbDataAdapter("SELECT * FROM Время WHERE (flag=True)", myOleDbConnection);
        myDataSet = new DataSet("Время");
        myDataAdapter.Fill(myDataSet, "Время");
        myDataAdapter.SelectCommand.Connection.Close();

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Клиенты", myOleDbConnection);
        myDataAdapter.SelectCommand.Connection.Open();
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Клиенты");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Персонал", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Персонал");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Процедуры", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Процедуры");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Специальности", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Специальности");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM spa_процедуры", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "spa_процедуры");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Расписание");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Расписание2");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Расписание1");

        myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Расписание", myOleDbConnection);
        myDataAdapter.SelectCommand.ExecuteNonQuery();
        myDataAdapter.Fill(myDataSet, "Расписание3");

        myDataAdapter.SelectCommand.Connection.Close();

        this.dataGridView1.DataSource = myDataSet.Tables["Персонал"].DefaultView;
        this.dataGridView1.Columns["ID_Персонала"].Visible = false;
    }

  

    private void groupBox1_Enter(object sender, EventArgs e)
    {

    }

  
    private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
    {

    }

    private void button15_Click_1(object sender, EventArgs e)
    {
      button15.Visible = false;
      groupBox4.Visible = true;
      button10.Visible = false;
      button3.Visible = false;
    }

    private void button16_Click(object sender, EventArgs e)
    {
      button15.Visible = true; ;
      groupBox4.Visible = false;
      button10.Visible =true;
      button3.Visible = true;
    }

    private void label28_Click(object sender, EventArgs e)
    {

    }

    private void label27_Click(object sender, EventArgs e)
    {

    }

    private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
    {
     if (comboBox11.SelectedIndex != -1)
     {
      if (comboBox12.Text != "23")
      {
       comboBox10.SelectedIndex = comboBox12.SelectedIndex + 1;
       comboBox9.SelectedIndex = 0;
      }
      else
      {
       comboBox10.SelectedIndex = 0;
       comboBox9.SelectedIndex = 0;
      }
     }
    }

    private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
    {
     if (comboBox12.SelectedIndex != -1)
     {
         string com = " ", Mytime = " ";
         int hour = 0, minutes, raznost = 0;
         OleDbDataReader dr;    
        OleDbConnection myConn = new  OleDbConnection(connectionString);
        
         myConn.Open();
         
         if (comboBox12.Text != "23")
          {
           //comboBox10.SelectedIndex = comboBox12.SelectedIndex + 1;
           //comboBox9.SelectedIndex = 0;
          }
          else
          {
           comboBox10.SelectedIndex = 0;
           comboBox9.SelectedIndex = 0;
          }
         if (radioButton1.Checked == true)
             com = "SELECT Продолжительность FROM Процедуры WHERE (Название = '" + comboBox5.Text + "')";
         if (radioButton2.Checked == true)
             com = "SELECT Продолжительность FROM spa-процедуры WHERE (Название = '" + comboBox7.Text + "')";
         OleDbCommand cmd = new OleDbCommand(com, myConn);
         Mytime = cmd.ExecuteScalar().ToString();
         int newTime = Convert.ToInt32(Mytime);
        // MessageBox.Show(comboBox11.Text);
         if (Convert.ToInt32(comboBox11.Text) + newTime < 60)
         {
             comboBox10.SelectedIndex = comboBox12.SelectedIndex;
             comboBox9.Text = Convert.ToString (Convert.ToInt32(comboBox11.Text) + newTime);
         }
         else if (Convert.ToInt32(comboBox11.Text) + newTime == 60)
         {
             comboBox10.SelectedIndex = comboBox12.SelectedIndex + 1;
             comboBox9.SelectedIndex = 0;
         }
         else if (Convert.ToInt32(comboBox11.Text) + newTime > 60)
         {             
             raznost = Convert.ToInt32(comboBox11.Text) + newTime;
             while (raznost >= 60)
             {
                 raznost = raznost - 60;
                 hour++;
             }
             comboBox10.SelectedIndex = comboBox12.SelectedIndex + hour;
             comboBox9.Text =  Convert.ToString(raznost);
         }
         MessageBox.Show("H = " + Convert.ToString(hour) + " M = " + Convert.ToString(raznost));
         myConn.Close();

     }
    }

    private void tabPage1_Click(object sender, EventArgs e)
    {

    }

    private void comboBox6_EnabledChanged(object sender, EventArgs e)
    {
        if (radioButton1.Checked == false && radioButton2.Checked == false)
            return;
        string com = " ", com1 = " ", pers = " ";
        OleDbDataReader dr;
       // myOleDbConnection.Close();
        OleDbConnection myConn = new  OleDbConnection(connectionString);
        myConn.Open();
        if (comboBox6.Enabled == true)
        {
            pers = comboBox6.Text;
            if (radioButton1.Checked == true)
            {
                //com = "SELECT Фамилия FROM Персонал WHERE (Not Специальность='СПА-мастер')";                
                
                com1 = "SELECT Процедуры.Название, Процедуры.Продолжительность FROM Процедуры Процедуры INNER JOIN (Персонал INNER JOIN Персонал_Процедуры ON Персонал.ID_Персонала = Персонал_Процедуры.ID_Персонала) ON Процедуры.ID_Процедуры = Персонал_Процедуры.ID_Процедуры WHERE (Персонал.Фамилия = '" + pers + "')";
            
            }
            else if (radioButton2.Checked == true)
            {
                //comboBox6.DataSource = myDataSet.Tables["SPA_Персонал"].DefaultView;
                //comboBox6.DisplayMember = "Фамилия";
                //com = "SELECT Фамилия FROM Персонал  WHERE (Специальность='СПА-мастер')";  // SELECT Название FROM spa_процедуры;
                com1 = "SELECT spa_процедуры.Название, spa_процедуры.Продолжительность FROM Персонал INNER JOIN spa_процедуры ON Персонал.ID_Персонала = spa_процедуры.ID_Персонала WHERE ((Персонал.Фамилия)='" + pers + "')";
              // SELECT spa_процедуры.Название, spa_процедуры.Продолжительность FROM Персонал INNER JOIN spa_процедуры ON Персонал.ID_Персонала = spa_процедуры.ID_Персонала WHERE ((Персонал.Фамилия)="Моренков");
            }
            OleDbCommand cmd = new OleDbCommand(com1, myConn);

            dr = cmd.ExecuteReader();
            if (radioButton1.Checked == true)
            {
                comboBox5.Items.Clear();
                while (dr.Read())
                    comboBox5.Items.Add(dr[0].ToString());
                comboBox5.SelectedIndex = 0;
            }
            if (radioButton2.Checked == true)
            {
                comboBox7.Items.Clear();
                while (dr.Read())
                    comboBox7.Items.Add(dr[0].ToString());
                comboBox7.SelectedIndex = 0;
            }
            dr.Close();
            //myOleDbConnection.Close();
           // myOleDbConnection.Open();
           // cmd = new OleDbCommand(com1, myOleDbConnection);
           /* dr = cmd.ExecuteReader();
            comboBox6.Items.Clear();
            while (dr.Read())
                comboBox6.Items.Add(dr[0].ToString());
            comboBox6.SelectedIndex = 0;
            dr.Close();
            myOleDbConnection.Close();*/
        }
        myConn.Close();
    }

    private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
    {
        comboBox6_EnabledChanged (sender, e);
    }

    private void groupBox2_Enter(object sender, EventArgs e)
    {

    }

    private void label24_Click(object sender, EventArgs e)
    {

    }

    private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
    {

    }

  }
}
