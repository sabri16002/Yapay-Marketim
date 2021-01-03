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
    public partial class Musteri : Form
    {
        DataTable dt_Musteri;
        BindingSource bs_Musteri;
        OleDbDataAdapter a1;
        OleDbConnection con;
        string sorgu;
        OleDbCommand komut = new OleDbCommand();
        public Musteri()
        {
            InitializeComponent();
        }
        void nesneler()
        {
            textBox1.DataBindings.Clear();
            textBox2.DataBindings.Clear();
            textBox3.DataBindings.Clear();

            textBox1.DataBindings.Add("text", bs_Musteri, "Musteri_no");
            textBox2.DataBindings.Add("text", bs_Musteri, "Ad");
            textBox3.DataBindings.Add("text", bs_Musteri, "Soyad");
        }
        void listele()
        {
            dt_Musteri = new DataTable();
            a1 = new OleDbDataAdapter(sorgu, con);
            a1.Fill(dt_Musteri);
            bs_Musteri = new BindingSource();
            bs_Musteri.DataSource = dt_Musteri;
            dataGridView1.DataSource = bs_Musteri;
            dataGridView1.RowHeadersWidth = 15;
            dataGridView1.Columns[0].Width = 80; dataGridView1.Columns[1].Width = 131;
            nesneler();
        }
        private void Musteri_Load(object sender, EventArgs e)
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
            sorgu = "Select *from Musteri";
            listele();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                komut.Connection = con;
                komut.CommandText = "insert into Musteri (Ad,Soyad,Tel) values ('" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "')";
                komut.ExecuteNonQuery();
                komut.Dispose();

                sorgu = "Select *from Musteri";
                MessageBox.Show("Müşteri Eklendi");
                listele();
            }
            catch (Exception hata)
            {
                MessageBox.Show(hata.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult cevap;
                cevap = MessageBox.Show("Müşteriyi silmek istediğinizden eminmisiniz", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes)
                {
                    komut.Connection = con;
                    komut.CommandText = "delete from Musteri where Musteri_no=" + dataGridView1.CurrentRow.Cells["Musteri_no"].Value + "";
                    komut.ExecuteNonQuery();
                    komut.Dispose();
                    sorgu = "select * from Musteri";
                    MessageBox.Show("Müşteri Silindi");
                    listele();
                }
            }
            catch (Exception hata)
            {
                MessageBox.Show(hata.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            sorgu = "Select * from Musteri Where Musteri_no like '" + textBox1.Text + "%'";
            listele();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            sorgu = "Select *from Musteri";
            listele();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            XlsFile excel = new XlsFile(true);
            excel.NewFile();

            excel.SetCellValue(1, 3, "Müşteri");
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
