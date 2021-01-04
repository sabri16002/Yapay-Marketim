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
using FlexCel.XlsAdapter;
using FlexCel.Core;

namespace Yapay_Marketim
{
    public partial class Satis : Form
    {
        DataTable dt_Satis, dt_Musteri, dt_Urun;
        BindingSource bs_Satis, bs_Musteri, bs_Urun;
        OleDbDataAdapter a1, a2, a3;
        OleDbConnection con;
        string sorgu;
        int mik;
        double fiy, tut;
        OleDbCommand komut = new OleDbCommand();
        public Satis()
        {
            InitializeComponent();
        }
        void listele()
        {
            dt_Satis = new DataTable();
            a1 = new OleDbDataAdapter(sorgu, con);
            a1.Fill(dt_Satis);
            bs_Satis = new BindingSource();
            bs_Satis.DataSource = dt_Satis;
            dataGridView1.DataSource = bs_Satis;
            dataGridView1.RowHeadersWidth = 15;
            dataGridView1.Columns[0].Width = 90; dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].Width = 100; dataGridView1.Columns[3].Width = 85;
            dataGridView1.Columns[4].Width = 50; dataGridView1.Columns[5].Width = 85;
            dataGridView1.Columns[6].Width = 100;
        }
        void listele2()
        {
            dt_Musteri = new DataTable();
            a2 = new OleDbDataAdapter(sorgu, con);
            a2.Fill(dt_Musteri);
            bs_Musteri = new BindingSource();
            bs_Musteri.DataSource = dt_Musteri;
        }

        void listele3()
        {
            dt_Urun = new DataTable();
            a3 = new OleDbDataAdapter(sorgu, con);
            a3.Fill(dt_Urun);
            bs_Urun = new BindingSource();
            bs_Urun.DataSource = dt_Urun;
            dataGridView2.DataSource = bs_Urun;
            dataGridView2.RowHeadersWidth = 15;
        }
        private void Satis_Load(object sender, EventArgs e)
        {
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox6.Enabled = false;
            dateTimePicker1.Value = DateTime.Today;
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
            sorgu = "Select *from Satis";
            listele();
            sorgu = "Select *from Urun";
            listele3();
        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            sorgu = "Select *from Musteri Where Musteri_no=" + comboBox1.Text + "";
            listele2();
            textBox1.Text = dt_Musteri.Rows[0]["Ad"].ToString();
            textBox2.Text = dt_Musteri.Rows[0]["Soyad"].ToString();
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            sorgu = "Select *from Musteri";
            listele2();
            comboBox1.Items.Clear();
            for (int k = 0; k <= dt_Musteri.Rows.Count - 1; k++)
            {
                comboBox1.Items.Add(dt_Musteri.Rows[k]["Musteri_no"].ToString());
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox3.Clear();
            textBox4.Clear();
            textBox6.Clear();
            sorgu = "Select *from Urun Where Urun_no='" + comboBox2.Text + "'";
            listele3();
            textBox3.Text = dt_Urun.Rows[0]["Urun"].ToString();
            textBox4.Text = dt_Urun.Rows[0]["Fiyat"].ToString();
        }

        private void comboBox2_Click(object sender, EventArgs e)
        {
            sorgu = "Select *from Urun";
            listele3();
            comboBox2.Items.Clear();
            for (int k = 0; k <= dt_Urun.Rows.Count - 1; k++)
            {
                comboBox2.Items.Add(dt_Urun.Rows[k]["Urun_no"].ToString());
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            textBox6.Clear();
            fiy = Convert.ToDouble(textBox4.Text);
            mik = Convert.ToInt32(textBox5.Text);
            tut = fiy * mik;
            textBox6.Text = tut.ToString();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                komut.Connection = con;
                komut.CommandText = "insert into Satis (Musteri_no,Ad,Soyad,Urun_no,Urun,Fiyat,Miktar,Tutar,Tarih) values (" +
              " '" + comboBox1.Text + "','" + textBox1.Text + "','" + textBox2.Text + "','" + comboBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + dateTimePicker1.Value.ToShortDateString() + "')";
                komut.ExecuteNonQuery();
                komut.Dispose();
                MessageBox.Show("Satış Yapıldı", textBox1.Text);
            }

            catch (Exception hata)
            {
                MessageBox.Show(hata.Message);
            }
            sorgu = "Select *from Satis";
            listele();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult cevap;
                cevap = MessageBox.Show("Kayıtı silmek istediğinizden emin misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes)
                {
                    komut.Connection = con;
                    komut.CommandText = "delete from Satis where Musteri_no=" + dataGridView1.CurrentRow.Cells["Musteri_no"].Value + "";
                    komut.ExecuteNonQuery();
                    komut.Dispose();
                    sorgu = "select * from Satis";
                    MessageBox.Show("Satış Kayıtı Silindi");
                    listele();
                }
            }
            catch (Exception hata)
            {
                MessageBox.Show(hata.Message);
            }
            sorgu = "Select *from Satis";
            listele();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            sorgu = "Select *from Satis Where Day(Tarih) = '" + comboBox4.Text + "'";
            listele();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            sorgu = "Select *from Satis Where Month(Tarih) = '" + comboBox3.Text + "'";
            listele();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            sorgu = "Select *from Satis Where Year(Tarih) = '" + textBox7.Text + "'";
            listele();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            sorgu = "Select *from Satis Where Day(Tarih) = '" + comboBox4.Text + "'";
            sorgu = "Select *from Satis Where Month(Tarih) = '" + comboBox3.Text + "'";
            sorgu = "Select *from Satis Where Year(Tarih) = '" + textBox7.Text + "'";
            listele();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            sorgu = "Select *from Satis";
            listele();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            XlsFile excel = new XlsFile(true);
            excel.NewFile();

            excel.SetCellValue(1, 5, " Şatış Tablosu");
            int ek = 3;
            for (int i = 1; i <= dataGridView1.ColumnCount; i++)
            {
                excel.SetCellValue(3, i, dataGridView1.Columns[i - 1].Name);
            }

            for (int i = 1; i <= dataGridView1.RowCount - 1; i++)
            {

                for (int k = 1; k <= dataGridView1.ColumnCount; k++)
                {
                    excel.SetCellValue(i + ek, k, dataGridView1[k - 1, i - 1].Value.ToString());

                }
            }
            saveFileDialog1.Filter = "*.xlsx|*.xlsx";
            saveFileDialog1.ShowDialog();
            string yol = saveFileDialog1.FileName;

            try
            {
                excel.Save("" + yol + "");
            }
            catch
            {
                MessageBox.Show("Aktarma Başarısız.."); return;
            }
            MessageBox.Show("Excele Aktarma İşlemi Bitti.");
        }

    }
}
