using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;
using System.Net;
using System.Net.Mail;


using System.Diagnostics;

using System.Web ;



namespace SPA
{
    
    public partial class Login : Form
    {
        BossAdd BossAdd;
         int a;
  //    Login2 Login2;
        public Login()
        {
          this.Width=315;
          this.Height = 132;
            InitializeComponent();
  
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Login_Load(object sender, EventArgs e)
        {
          this.Width = 294;
          this.Height = 132;
          this.Opacity = 0.9;
          textBox3.Enabled = false;
          textBox3.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "Password")
            {
              textBox1.Clear();
              textBox1.Enabled = false;
              textBox1.Visible = false;
              textBox3.Visible = true;
              textBox3.Enabled = true;
              MessageBox.Show("Сейчас Вам на почту придет письмо с кодом доступа, введите его", "Внимание!");
              this.Width=605;
              this.Height = 132;
              button1.Enabled = false;
              //if (textBox3.Text == a.ToString())
              //{
              //  BossAdd = new BossAdd();
              //  BossAdd.Owner = this;
              //  BossAdd.ShowDialog();
              //  //this.Opacity = 1.0;
              //  this.Hide();
              //}
              //else MessageBox.Show("Неверный код доступа!");
                   
            }
            else
            {
                MessageBox.Show("Неверный пароль!");
            }
            
        }
        private void auth()
        {
          MailMessage message;
          SmtpClient client;
          System.Random random = new System.Random();
          int random_value = random.Next();
          a = random_value;


          message = new System.Net.Mail.MailMessage(
               "MailTo<mail.to.spaforvip@gmail.com>",
               "mail.spaforvip@gmail.com",
               "Код доступа",
              random_value.ToString());

          
      client = new SmtpClient("smtp.gmail.com", 587)
      {
        Credentials = new NetworkCredential("mail.spaforvip@gmail.com", textBox2.Text),
        EnableSsl = true
      };

      message.BodyEncoding = System.Text.Encoding.UTF8;
      message.IsBodyHtml = true;

      
      //message.
      message.Attachments.Add(new Attachment("qr_code.png"));
      try
            {
                client.Send(message);
                message.Attachments.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.ToString());
                message.Attachments.Dispose();
                return;
            }

            return;

    }

      private void textBox2_TextChanged(object sender, EventArgs e)
      {
      
      }

        
        private void groupBox1_Enter(object sender, EventArgs e)
        {
          
        }

        private void groupBox1_Enter_1(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
          this.Width = 294;
          this.Height = 132;
        }

        private void button4_Click(object sender, EventArgs e)
        {
          if (textBox2.Text != null)
          {
            button4.Enabled = false;
            auth();

            //SHDocVw.InternetExplorer iexplorer = new SHDocVw.InternetExplorer();

            Process prc = new Process();
            //prc.StartInfo.FileName = @"C:\Program Files\Internet Explorer\iexplore.exe";
           System.Diagnostics.Process.Start("https://gmail.com");         

            this.Width = 967;
            this.Height = 132;
          }
          else MessageBox.Show("Введите пароль от почты", "Внимание!");
        }

        private void button5_Click(object sender, EventArgs e)
        {
          button4.Enabled = true;
          this.Width = 605;
          this.Height = 132;
        }

        private void button6_Click(object sender, EventArgs e)
        {
          if (textBox4.Text == a.ToString())
          {
           // System.Diagnostics.Process.
            BossAdd = new BossAdd();
            BossAdd.Owner = this;
            BossAdd.ShowDialog();

            //this.Opacity = 1.0;
            this.Hide();
          }
        }
    }
}
