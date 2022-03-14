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
    public partial class kelimeEkle : Form
    {
        public kelimeEkle()
        {
            InitializeComponent();
            
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

        private void button1_Click(object sender, EventArgs e)
        {
           
            int nRowIndex = dataGridView1.Rows.Count;
            int b = nRowIndex + 1;
            OleDbCommand komut = new OleDbCommand();
            baglanti.Open();
            komut.Connection = baglanti;
            string sql = "insert into [Sayfa1$] (Numara,Türkçe,English) values('" + b + "','" + textBox1.Text + "','" + textBox2.Text + "')";
            komut.CommandText = sql;
            komut.ExecuteNonQuery();
            baglanti.Close();
            doldur();

            textBox1.Clear();
            textBox1.Text = "";
            textBox2.Clear();
            textBox2.Text = "";



        }

        private void button2_Click(object sender, EventArgs e)
        {

            doldur();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                button1_Click(this, new EventArgs());
            }
        }

        private void kelimeEkle_Load(object sender, EventArgs e)
        {
            doldur();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int nRowIndex = dataGridView1.Rows.Count;
            int b = nRowIndex + 1;
            OleDbCommand komut = new OleDbCommand();
            baglanti.Open();
            komut.Connection = baglanti;
            string sql = "insert into [Sayfa1$] (Numara,Türkçe,Türkçe2,English) values('" + b + "','" + textBox5.Text + "','" + textBox4.Text + "','" + textBox3.Text + "')";
            komut.CommandText = sql;
            komut.ExecuteNonQuery();
            baglanti.Close();
            doldur();

            textBox3.Clear();
            textBox3.Text = "";
            textBox4.Clear();
            textBox4.Text = "";
            textBox5.Clear();
            textBox5.Text = "";
       
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                button3_Click(this, new EventArgs());
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.TextLength > 0)
            {
                textBox1.Text = char.ToUpper(textBox1.Text[0]).ToString() + textBox1.Text.Substring(1);
                textBox1.SelectionStart = textBox1.TextLength;
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

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (textBox5.TextLength > 0)
            {
                textBox5.Text = char.ToUpper(textBox5.Text[0]).ToString() + textBox5.Text.Substring(1);
                textBox5.SelectionStart = textBox5.TextLength;
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.TextLength > 0)
            {
                textBox4.Text = char.ToUpper(textBox4.Text[0]).ToString() + textBox4.Text.Substring(1);
                textBox4.SelectionStart = textBox4.TextLength;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.TextLength > 0)
            {
                textBox3.Text = char.ToUpper(textBox3.Text[0]).ToString() + textBox3.Text.Substring(1);
                textBox3.SelectionStart = textBox3.TextLength;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
          

        }
    }
}
