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
using System.IO;

namespace Sozluk
{
    public partial class TrToEng : Form
    {
        public TrToEng()
        {
            InitializeComponent();
        }

        private void TrToEng_Load(object sender, EventArgs e)
        {
            doldur();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Form1 ff = new Form1();
            ff.Show();
            this.Close();
        }
        private void button1_Click(object sender, EventArgs e)
        {

            int a = int.Parse(textBox3.Text);
            
            if (textBox2.Text == dataGridView1.Rows[a].Cells[2].Value.ToString())
            {
                label1.Text = "Doğru";
                label6.Text= dataGridView1.Rows[a].Cells[2].Value.ToString();
                label4.Text = "";
                label5.Text = "";
            }
            else
            {
                label5.Text = "Doğrusu:";
                label1.Text = "Yanlış";
                label4.Text = dataGridView1.Rows[a].Cells[2].Value.ToString();
                label6.Text = "...";
            }
        }
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\English\Kelime Çalışması.xls; Extended Properties='Excel 12.0 xml;'");
        //HDR=YES;

        public void doldur()
        {
            baglanti.Open();
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt.DefaultView;
            baglanti.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            textBox2.Text = "";
            Random rastgele = new Random();
            int satir = rastgele.Next(0, dataGridView1.Rows.Count);

            
            textBox3.Text = satir.ToString();

            var rows = dataGridView1.Rows;
            if (rows[satir].Cells[1] != null)
            {
                textBox1.Text = dataGridView1.Rows[satir].Cells[1].Value.ToString();

            }

        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                button1_Click(this, new EventArgs());
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.TextLength > 0)
            {
                textBox2.Text = char.ToUpper(textBox2.Text[0]).ToString() + textBox2.Text.Substring(1);
                textBox2.SelectionStart = textBox2.TextLength;
            }
        }

        private void TrToEng_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Left)
            {
                //button2.Select();
                button2.PerformClick();
            }
        }
    }
}
