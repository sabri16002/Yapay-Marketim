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

namespace Yapay_Marketim
{
    public partial class Login : Form
    {
        OleDbConnection con;
        OleDbDataReader gir, yonetici;
        string sorgu;
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

            con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Yapay Marketim.mdb");
            con.Open();
            Ana ana = new Ana();
            sorgu = "SELECT * FROM login";
            OleDbCommand kom = new OleDbCommand(sorgu, con);
            kom.Connection = con;
            kom.CommandText = "SELECT * FROM Giris where Ad='" + textBox1.Text + "' AND Sifre='" + textBox2.Text + "'";
            gir = kom.ExecuteReader();
            if (gir.Read())
            {
                gir.Close();
                kom.CommandText = "SELECT * FROM Giris where Ad= '" + textBox1.Text + "'and Yonet=" + 0 + "";
                yonetici = kom.ExecuteReader();
                if (yonetici.Read())
                {
                    ana.toolStripButton2.Enabled = false;
                    ana.toolStripButton3.Enabled = false;
                    ana.toolStripButton4.Enabled = false;
                }
                MessageBox.Show("Tebrikler! Başarılı bir şekilde giriş yaptınız.");
                ana.Show();
                this.Visible = false;
            }
            else
            {
                textBox1.Clear();
                textBox2.Clear();
                MessageBox.Show("Kullanıcı adını veya şifre hatalı kontrol ediniz.");
            }
            con.Close();
        }

    }
}
