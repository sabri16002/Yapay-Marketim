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
    public partial class Ana : Form
    {
        public Ana()
        {
            InitializeComponent();
        }
        private void Ana_Load(object sender, EventArgs e)
        {
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Satis sat = new Satis();
            sat.MdiParent = this;
            sat.Show();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Urunler ur = new Urunler();
            ur.MdiParent = this;
            ur.Show();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            Musteri mus = new Musteri();
            mus.MdiParent = this;
            mus.Show();
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            Ayarlar ayar = new Ayarlar();
            ayar.MdiParent = this;
            ayar.Show();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Normal)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
