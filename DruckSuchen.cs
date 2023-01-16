using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Querdruck
{
    public partial class DruckSuchen : Form
    {
        private int index;
        private readonly string Datei;
        public DruckSuchen(string Datei)
        {
            InitializeComponent();
            this.Datei = Datei;
        }
        private void DruckSuchen_Load(object sender, EventArgs e)
        {
            DataTable dataTable = new DataTable();
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "select * from " + Datei;
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();
                dataTable.Load(rdr);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dataGridView1.DataSource = dataTable;
            try
            {
                dataGridView1.Columns[0].Visible = false;
                for (int i = 16; i < 53; i++)
                {
                    dataGridView1.Columns[i].Visible = false;
                }
                for (int i = 1; i < 15; i++)
                {
                    dataGridView1.Columns[i].Width = 200;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            DruckSuchenOK.Select();
        }

        private void DruckSuchenOK_Click(object sender, EventArgs e)
        {
            try
            {
                DataGridViewRow SelectedRow = dataGridView1.Rows[index];
                Form1.stzNr = SelectedRow.Cells[0].Value.ToString();
                Form1.z1 = SelectedRow.Cells[1].Value.ToString();
                Form1.z2 = SelectedRow.Cells[2].Value.ToString();
                Form1.z3 = SelectedRow.Cells[3].Value.ToString();
                Form1.z4 = SelectedRow.Cells[4].Value.ToString();
                Form1.z5 = SelectedRow.Cells[5].Value.ToString();
                Form1.z6 = SelectedRow.Cells[6].Value.ToString();
                Form1.z7 = SelectedRow.Cells[7].Value.ToString();
                Form1.z8 = SelectedRow.Cells[8].Value.ToString();
                Form1.z9 = SelectedRow.Cells[9].Value.ToString();
                Form1.z10 = SelectedRow.Cells[10].Value.ToString();
                Form1.z11 = SelectedRow.Cells[11].Value.ToString();
                Form1.z12 = SelectedRow.Cells[12].Value.ToString();
                Form1.z13 = SelectedRow.Cells[13].Value.ToString();
                Form1.z14 = SelectedRow.Cells[14].Value.ToString();
                Form1.z15 = SelectedRow.Cells[15].Value.ToString();
                Form1.srft = SelectedRow.Cells[16].Value.ToString();
                Form1.h1 = SelectedRow.Cells[17].Value.ToString();
                Form1.h2 = SelectedRow.Cells[18].Value.ToString();
                Form1.h3 = SelectedRow.Cells[19].Value.ToString();
                Form1.h4 = SelectedRow.Cells[20].Value.ToString();
                Form1.h5 = SelectedRow.Cells[21].Value.ToString();
                Form1.h6 = SelectedRow.Cells[22].Value.ToString();
                Form1.h7 = SelectedRow.Cells[23].Value.ToString();
                Form1.h8 = SelectedRow.Cells[24].Value.ToString();
                Form1.h9 = SelectedRow.Cells[25].Value.ToString();
                Form1.h10 = SelectedRow.Cells[26].Value.ToString();
                Form1.h11 = SelectedRow.Cells[27].Value.ToString();
                Form1.h12 = SelectedRow.Cells[28].Value.ToString();
                Form1.h13 = SelectedRow.Cells[29].Value.ToString();
                Form1.h14 = SelectedRow.Cells[30].Value.ToString();
                Form1.h15 = SelectedRow.Cells[31].Value.ToString();
                Form1.s1 = SelectedRow.Cells[32].Value.ToString();
                Form1.s2 = SelectedRow.Cells[33].Value.ToString();
                Form1.s3 = SelectedRow.Cells[34].Value.ToString();
                Form1.s4 = SelectedRow.Cells[35].Value.ToString();
                Form1.s5 = SelectedRow.Cells[36].Value.ToString();
                Form1.s6 = SelectedRow.Cells[37].Value.ToString();
                Form1.s7 = SelectedRow.Cells[38].Value.ToString();
                Form1.s8 = SelectedRow.Cells[39].Value.ToString();
                Form1.s9 = SelectedRow.Cells[40].Value.ToString();
                Form1.s10 = SelectedRow.Cells[41].Value.ToString();
                Form1.s11 = SelectedRow.Cells[42].Value.ToString();
                Form1.s12 = SelectedRow.Cells[43].Value.ToString();
                Form1.s13 = SelectedRow.Cells[44].Value.ToString();
                Form1.s14 = SelectedRow.Cells[45].Value.ToString();
                Form1.s15 = SelectedRow.Cells[46].Value.ToString();
                Form1.AVA = SelectedRow.Cells[47].Value.ToString();
                Form1.frb = SelectedRow.Cells[48].Value.ToString();
                Form1.BdNr = SelectedRow.Cells[49].Value.ToString();
                Form1.brte = SelectedRow.Cells[50].Value.ToString();
                Form1.gedd = SelectedRow.Cells[51].Value.ToString();
                Form1.Dtm = SelectedRow.Cells[52].Value.ToString();
                this.Close();
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Bitte Zeile wählen oder auf Abbrech kliclen!");
            }
        }

        private void DruckSuchenCancel_Click(object sender, EventArgs e)
        {
            Application.OpenForms[0].Show();
            this.Close();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            index = e.RowIndex;
        }

        private void SearchBox_TextChanged(object sender, EventArgs e)
        {
            DataTable dataTable = new DataTable();
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            string sql;

            try
            {
                conn.Open();
                if (string.IsNullOrEmpty(SearchBox.Text))
                {
                    sql = "select * from " + Datei;
                }
                else
                {
                    sql = string.Format("select * from " + Datei + " where Zeile1 like '{0}%' or " +
                    "Zeile2 like '{0}%' or Zeile3 like '{0}%' or Zeile4 like '{0}%' or " +
                    "Zeile5 like '{0}%' or Zeile6 like '{0}%' or Zeile7 like '{0}%' or " +
                    "Zeile8 like '{0}%' or Zeile9 like '{0}%' or Zeile10 like '{0}%' or " +
                    "Zeile11 like '{0}%' or Zeile12 like '{0}%' or Zeile13 like '{0}%' or " +
                    "Zeile14 like '{0}%' or Zeile15 like '{0}%';", SearchBox.Text);
                }
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();
                dataTable.Load(rdr);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dataGridView1.DataSource = dataTable;
            try
            {
                dataGridView1.Columns[0].Visible = false;
                for (int i = 16; i < 53; i++)
                {
                    dataGridView1.Columns[i].Visible = false;
                }
                for (int i = 1; i < 15; i++)
                {
                    dataGridView1.Columns[i].Width = 200;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            dataGridView1.ClearSelection();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DruckSuchenOK.PerformClick();
        }
    }
}
