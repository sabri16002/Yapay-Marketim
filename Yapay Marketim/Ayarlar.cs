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
    public partial class Ayarlar : Form
    {
        DataTable dt_Ayar;
        BindingSource bs_Ayar;
        OleDbDataAdapter a1;
        OleDbConnection con;
        string sorgu;
        int yon = 1, yon2 = 0;
        OleDbCommand komut = new OleDbCommand();
        public Ayarlar()
        {
            InitializeComponent();
        }
        void nesneler()
        {
            textBox1.DataBindings.Clear();
            textBox1.DataBindings.Add("text", bs_Ayar, "Ad");
        }
        void listele()
        {
            dt_Ayar = new DataTable();
            a1 = new OleDbDataAdapter(sorgu, con);
            a1.Fill(dt_Ayar);
            bs_Ayar = new BindingSource();
            bs_Ayar.DataSource = dt_Ayar;
            dataGridView1.DataSource = bs_Ayar;
            dataGridView1.RowHeadersWidth = 15;
            nesneler();
        }
        private void Ayarlar_Load(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Yapay Marketim.mdb");
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
            }
            catch
            {
                MessageBox.Show("veritabanı ile bağlantı sağlanamadı");
            }
            sorgu = "Select *from Giris";
            listele();
        }
        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                komut.Connection = con;
                komut.CommandText = "insert into Giris (Ad,Sifre,Yonet) values ('" + textBox1.Text + "','" + textBox2.Text + "','" + yon2 + "')";
                komut.ExecuteNonQuery();
                komut.Dispose();

                sorgu = "Select *from Giris";
                MessageBox.Show("Kullancı Eklendi");
                listele();
            }

            catch (Exception hata)
            {
                MessageBox.Show(hata.Message);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {

            DialogResult sonuc = MessageBox.Show("Yetki Vereyim mi ? ", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (sonuc == DialogResult.No)
                return;

            try
            {
                komut.Connection = con;
                komut.CommandText = "UPDATE Giris SET Yonet='" + yon + "'Where Ad= '" + dataGridView1.CurrentRow.Cells["Ad"].Value + "'";
                komut.ExecuteNonQuery();
                komut.Dispose();

                sorgu = "Select *from Giris order by Ad Asc";
                listele();
            }

            catch (Exception hata)
            {
                MessageBox.Show(hata.Message);
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult sonuc = MessageBox.Show("Yetkiyi Geri Alayım mı ? ", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (sonuc == DialogResult.No)
                return;

            try
            {
                komut.Connection = con;
                komut.CommandText = "UPDATE Giris SET Yonet='" + yon2 + "'Where Ad= '" + dataGridView1.CurrentRow.Cells["Ad"].Value + "'";
                komut.ExecuteNonQuery();
                komut.Dispose();

                sorgu = "Select *from Giris order by Ad Asc";
                listele();
            }

            catch (Exception hata)
            {
                MessageBox.Show(hata.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            sorgu = "Select * from Giris Where Ad like '" + textBox1.Text + "%'";
            listele();
        }
    }
}
