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
    public partial class Urunler : Form
    {
        DataTable dt_urun;
        BindingSource bs_urun;
        OleDbDataAdapter a1;
        OleDbConnection con;
        string sorgu;
        OleDbCommand komut = new OleDbCommand();

        public Urunler()
        {
            InitializeComponent();
        }
        void nesneler()
        {
            textBox2.DataBindings.Clear();
            textBox3.DataBindings.Clear();
            textBox2.DataBindings.Add("text", bs_urun, "Urun");
            textBox3.DataBindings.Add("text", bs_urun, "Fiyat");
        }
        public void listele()
        {
            dt_urun = new DataTable();
            a1 = new OleDbDataAdapter(sorgu, con);
            a1.Fill(dt_urun);
            bs_urun = new BindingSource();
            bs_urun.DataSource = dt_urun;
            dataGridView1.DataSource = bs_urun;
            nesneler();
        }
        private void Urunler_Load(object sender, EventArgs e)
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
            sorgu = "Select *from Urun";
            listele();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                komut.Connection = con;
                komut.CommandText = "Insert into Urun(Urun_no,Urun,Fiyat) values (" +
              " '" + textBox1.Text + "','" + textBox2.Text + "','" + Convert.ToDouble(textBox3.Text) + "')";
                komut.ExecuteNonQuery();
                komut.Dispose();

                sorgu = "Select * from Urun"; 
                listele();
                MessageBox.Show("KAYIT BAŞARILI");
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
                cevap = MessageBox.Show("Ürünü silmek istediğinizden eminmisiniz", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes)
                {
                    komut.Connection = con;
                    komut.CommandText = "delete from Urun where Urun_no='" + dataGridView1.CurrentRow.Cells["Urun_no"].Value + "' ";
                    komut.ExecuteNonQuery();
                    komut.Dispose();
                    sorgu = "select * from Urun";
                    MessageBox.Show("Ürün Silindi");
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

            sorgu = "Select * from Urun Where Urun_no like '" + textBox1.Text + "%'";
            listele();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult sonuc = MessageBox.Show("Ürünü Güncellemek İstiyormusunuz ? ", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (sonuc == DialogResult.No)
                return;

            try
            {
                komut.Connection = con;
                komut.CommandText = "UPDATE Urun SET Urun_no= '" + textBox1.Text + "', Urun = '" + textBox2.Text + "',Fiyat ='" + textBox3.Text + "'Where Urun_no= '" + dataGridView1.CurrentRow.Cells["Urun_no"].Value + "'";
                komut.ExecuteNonQuery();
                komut.Dispose();

                sorgu = "Select *from Urun order by Urun Asc";
                listele();
            }

            catch (Exception hata)
            {
                MessageBox.Show(hata.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            XlsFile excel = new XlsFile(true);
            excel.NewFile();

            excel.SetCellValue(1, 2, " Ürünler");
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
