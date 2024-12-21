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

namespace Excel_Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        OleDbConnection bgl = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Koray\Desktop\orn20.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES;'");


        void listele()
        {
            bgl.Open();
            OleDbDataAdapter da = new OleDbDataAdapter("Select * From [Sayfa1$]", bgl);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            bgl.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listele();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bgl.Open();
            OleDbCommand cmd = new OleDbCommand("Insert Into [Sayfa1$] (Saat, Ders) Values (@p1, @p2)", bgl);
            cmd.Parameters.AddWithValue("@p1", textBox1.Text);
            cmd.Parameters.AddWithValue("@p2", textBox2.Text);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Yeni Ders Bilgisi Eklendi");
            listele();
        }
    }
}
