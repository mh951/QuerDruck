using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace Querdruck
{
    public partial class Form1 : Form
    {
        public static bool Bearbeiten_Mode = true;
        public static bool Druck_Mode = false;
        public static int letzte_Zeile = 1;
        public int Schriftgröße = -1;
        public int Seite2 = 0;
        public static string z1, z2, z3, z4, z5, z6, z7, z8, z9, z10, z11, z12, z13, z14, z15,
                             h1, h2, h3, h4, h5, h6, h7, h8, h9, h10, h11, h12, h13, h14, h15,
                             s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15,
                             stzNr, BdNr, brte, AVA, srft, frb, gedd, Dtm;
        Dictionary<TextBox, string> GoNull = new Dictionary<TextBox, string> { };
        public string AktuellDruck;
        public string AktuellDatei = "Druckdatei";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            BearbeitenMode();
        }
        
        // Wenn man auf eine Zeile ist und auf Enter kliclt, dann fokusieren auf nächsten Zeile 
        private void Zeile1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile2.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile3.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile4.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile5.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile6.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile7.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile7_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile8.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile8_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile9.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile9_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile10.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile10_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile11.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile11_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile12.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile12_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile13.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile13_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile14.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void Zeile14_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Zeile15.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
        }
       
        // Nur Nummer in "Sperren" erlauben
        private void Sperren1_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren2_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren3_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren4_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren5_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren6_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren7_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren8_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren9_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren10_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren11_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren12_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren13_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren14_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Sperren15_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        // Nur Nummer in "Länge" erlauben
        private void Laenge1_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge2_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge3_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge4_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge5_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge6_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge7_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge8_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge9_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge10_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge11_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge12_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge13_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge14_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Laenge15_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        // Nur Nummer in "Höhe" erlauben
        private void Hoehe1_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe2_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe3_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe4_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe5_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe6_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe7_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe8_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe9_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe10_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe11_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe12_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe13_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe14_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        private void Hoehe15_KeyPress(object sender, KeyPressEventArgs e)
        {
            {
                if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        // Überprüfen welche Zeile die aktuelle ist, um aus "Vorschläge" einen Vorschlag zu bekommen
        private void Zeile1_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 1;
        }

        private void Zeile2_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 2;
        }

        private void Zeile3_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 3;
        }

        private void Zeile4_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 4;
        }

        private void Zeile5_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 5;
        }

        private void Zeile6_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 6;
        }

        private void Zeile7_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 7;
        }

        private void Zeile8_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 8;
        }

        private void Zeile9_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 9;
        }

        private void Zeile10_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 10;
        }

        private void Zeile11_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 11;
        }

        private void Zeile12_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 12;
        }

        private void Zeile13_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 13;
        }

        private void Zeile14_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 14;
        }

        private void Zeile15_MouseClick(object sender, MouseEventArgs e)
        {
            if (SchriftGröße.Text == "")
            {
                MessageBox.Show("Bitte die Schriftgröße wählen");
                label5.Focus();
                return;
            }
            letzte_Zeile = 15;
        }

        // Vorschläge zu den Zeilen schicken + Auf eine Zeile ohne Schriftgröße zu schreiben verhindern
        private void Vorschläge_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (letzte_Zeile == 1)
            {
                Zeile1.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 2)
            {
                Zeile2.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 3)
            {
                Zeile3.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 4)
            {
                Zeile4.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 5)
            {
                Zeile5.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 6)
            {
                Zeile6.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 7)
            {
                Zeile7.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 8)
            {
                Zeile8.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 9)
            {
                Zeile9.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 10)
            {
                Zeile10.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 11)
            {
                Zeile11.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 12)
            {
                Zeile12.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 13)
            {
                Zeile13.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 14)
            {
                Zeile14.Text = Vorschläge.Text;
            }
            else if (letzte_Zeile == 15)
            {
                Zeile15.Text = Vorschläge.Text;
            }
        }

        // "Länge" berechnen
        private void SchritGrößeÄndernFürLänge(TextBox Zeile, TextBox Laenge)
        {
            if (Zeile.Text == "")
            {
                Laenge.Text = "";
                return;
            }
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string x = Zeile.Text;
                double y = 0;
                foreach (char c in x)
                {
                    string sql = "select Breite from Tabelle" + Schriftgröße + " where Zeichen = '" + c + "' COLLATE utf8mb4_bin;";
                    MySqlCommand cmd = new MySqlCommand(sql, conn);
                    MySqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        y += Int32.Parse(rdr[0].ToString());
                    }
                    rdr.Close();
                }
                y /= 10;
                Laenge.Text = y.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            conn.Close();
        }

        // Schicken die Zeielen um ihre Länge zu berechnen
        private void Zeile1_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile1, LaengeZeile1);
        }

        private void Zeile2_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile2, LaengeZeile2);
        }

        private void Zeile3_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile3, LaengeZeile3);
        }

        private void Zeile4_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile4, LaengeZeile4);
        }

        private void Zeile5_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile5, LaengeZeile5);
        }

        private void Zeile6_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile6, LaengeZeile6);
        }

        private void Zeile7_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile7, LaengeZeile7);
        }

        private void Zeile8_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile8, LaengeZeile8);
        }

        private void Zeile9_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile9, LaengeZeile9);
        }

        private void Zeile10_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile10, LaengeZeile10);
        }

        private void Zeile11_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile11, LaengeZeile11);
        }

        private void Zeile12_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile12, LaengeZeile12);
        }

        private void Zeile13_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile13, LaengeZeile13);
        }

        private void Zeile14_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile14, LaengeZeile14);
        }

        private void Zeile15_TextChanged(object sender, EventArgs e)
        {
            SchritGrößeÄndernFürLänge(Zeile15, LaengeZeile15);
        }

        private void FensterWechseln_Click(object sender, EventArgs e)
        {
            if (Bearbeiten_Mode == true)
            {
                DruckMode();
                Bearbeiten_Mode = false;
            } 
            else
            {
                BearbeitenMode();
                Bearbeiten_Mode = true;
            }
            
        }

        // "Breite" zum Index Nummer konvertieren (So ist es in der MS-Datenbank)
        private int BreiteToDatenBank(int Breite)
        {
            if (Breite == 75) return 1;
            else if (Breite == 100) return 2;
            else if (Breite == 125) return 3;
            else if (Breite == 150) return 4;
            else if (Breite == 175) return 5;
            else if (Breite == 200) return 6;
            else if (Breite == 225) return 7;
            else return -1;
        }

        // 'Ged'-Spalte erstellen (muss man mit 'x' erweitern, wenn es gedruckt ist)
        private string CreateGedQuerDruck()
        {
            string Ged = "Q";
            int SchriftInt = SchriftGröße.SelectedIndex + 1;
            string SchriftChar = SchriftInt.ToString();

            Ged = (Zeile1.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile2.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile3.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile4.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile5.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile6.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile7.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile8.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile9.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile10.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile11.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile12.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile13.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile14.Text == "") ? Ged + "0" : Ged + SchriftChar;
            Ged = (Zeile15.Text == "") ? Ged + "0" : Ged + SchriftChar;

            return Ged;
        }

        // Überprüfen ob jede Zeile ihre "Höhe" hat
        private bool CheckEveryZeileHasHöhe()
        {
            if (!string.IsNullOrEmpty(Zeile1.Text) && string.IsNullOrEmpty(Hoehe1.Text))
            {
                MessageBox.Show("Höhe Zeile 1 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile2.Text) && string.IsNullOrEmpty(Hoehe2.Text))
            {
                MessageBox.Show("Höhe Zeile 2 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile3.Text) && string.IsNullOrEmpty(Hoehe3.Text))
            {
                MessageBox.Show("Höhe Zeile 3 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile4.Text) && string.IsNullOrEmpty(Hoehe4.Text))
            {
                MessageBox.Show("Höhe Zeile 4 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile5.Text) && string.IsNullOrEmpty(Hoehe5.Text))
            {
                MessageBox.Show("Höhe Zeile 5 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile6.Text) && string.IsNullOrEmpty(Hoehe6.Text))
            {
                MessageBox.Show("Höhe Zeile 6 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile7.Text) && string.IsNullOrEmpty(Hoehe7.Text))
            {
                MessageBox.Show("Höhe Zeile 7 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile8.Text) && string.IsNullOrEmpty(Hoehe8.Text))
            {
                MessageBox.Show("Höhe Zeile 8 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile9.Text) && string.IsNullOrEmpty(Hoehe9.Text))
            {
                MessageBox.Show("Höhe Zeile 9 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile10.Text) && string.IsNullOrEmpty(Hoehe10.Text))
            {
                MessageBox.Show("Höhe Zeile 10 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile11.Text) && string.IsNullOrEmpty(Hoehe11.Text))
            {
                MessageBox.Show("Höhe Zeile 11 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile12.Text) && string.IsNullOrEmpty(Hoehe12.Text))
            {
                MessageBox.Show("Höhe Zeile 12 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile13.Text) && string.IsNullOrEmpty(Hoehe13.Text))
            {
                MessageBox.Show("Höhe Zeile 13 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile14.Text) && string.IsNullOrEmpty(Hoehe14.Text))
            {
                MessageBox.Show("Höhe Zeile 14 fehlt!");
                return false;
            }
            else if (!string.IsNullOrEmpty(Zeile15.Text) && string.IsNullOrEmpty(Hoehe15.Text))
            {
                MessageBox.Show("Höhe Zeile 15 fehlt!");
                return false;
            }
            else
            {
                return true;
            }
        }

        // Das Problem lösen, dass keine Null an die Datenbank zu senden und keine Leerzeilen zuzulassen
        private void CreateNullForDatenbank()
        {
            List<TextBox> textBoxes = new List<TextBox>
            {
                Zeile1, Zeile2, Zeile3, Zeile4, Zeile5,
                Zeile6, Zeile7, Zeile8, Zeile9, Zeile10,
                Zeile11, Zeile12, Zeile13, Zeile14, Zeile15,
                Hoehe1, Hoehe2, Hoehe3, Hoehe4, Hoehe5,
                Hoehe6, Hoehe7, Hoehe8, Hoehe9, Hoehe10,
                Hoehe11, Hoehe12, Hoehe13, Hoehe14, Hoehe15,
                Sperren1, Sperren2, Sperren3, Sperren4, Sperren5,
                Sperren6, Sperren7, Sperren8, Sperren9, Sperren10,
                Sperren11, Sperren12, Sperren13, Sperren14, Sperren15,
                BandNr
            };
            foreach (TextBox t in textBoxes)
            {
                if (t.Text == "")
                {
                    try
                    {
                        GoNull.Add(t, "null");
                    }
                    catch { GoNull[t] = "null"; }
                }
                else
                {
                    try
                    {
                        GoNull.Add(t, "\"" + t.Text + "\"");
                    }
                    catch { GoNull[t] = "\"" + t.Text + "\""; }
                }
            }
        }

        private void DruckenOrSpeichen_Click(object sender, EventArgs e)
        {
            if (Bearbeiten_Mode)
            {
                if (CheckEveryZeileHasHöhe() == false) { return; }
                CreateNullForDatenbank();
                try
                {
                    if (ArchivCheckBox.Checked && !DruckCheckBox.Checked)
                    {
                        DruckSpeichern("Archiv");
                        MessageBox.Show("Der Druck wurde in Archiv gespeichert");
                    }
                    else if (!ArchivCheckBox.Checked)
                    {
                        DruckSpeichern("Druckdatei");
                        MessageBox.Show("Der Druck wurde in Druckdatei gespeichert");
                    }
                    else if (ArchivCheckBox.Checked && DruckCheckBox.Checked)
                    {
                        DruckSpeichern("Archiv");
                        DruckSpeichern("Druckdatei");
                        MessageBox.Show("Der Druck wurde in Druckdatei und Archiv gespeichert");
                    }
                    EmptyTheFields();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else
            {
                // muss geändert werden
                return;
            }
        }

        // Speichern von Druck in Druckdatei oder Archiv oder beides
        private void DruckSpeichern(string SaveIn)
        {
            DateTime dt1 = DateTime.Now;
            string Datum = "\"" + dt1.ToString() + "\"";
            int breite = (Breite.Text == "") ? -1 : BreiteToDatenBank(Int32.Parse(Breite.Text));
            int farbe = (FarbeEingabe.Text == "") ? -1 : FarbeEingabe.SelectedIndex + 1;
            string Ged = "\"" + CreateGedQuerDruck() + "\"";
            string schrift = (SchriftGröße.SelectedIndex + 1).ToString();
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);

            try
            {
                conn.Open();
                string sql = "INSERT INTO " + SaveIn + "(Zeile1, Zeile2, Zeile3, Zeile4, Zeile5, Zeile6, Zeile7, Zeile8, Zeile9, " +
                    "Zeile10, Zeile11, Zeile12, Zeile13, Zeile14, Zeile15, Schrift, Höhe1, Höhe2, Höhe3, Höhe4, Höhe5, Höhe6, Höhe7," +
                    " Höhe8, Höhe9, Höhe10, Höhe11, Höhe12, Höhe13, Höhe14, Höhe15, Sperren1, Sperren2, Sperren3, Sperren4, Sperren5," +
                    " Sperren6, Sperren7, Sperren8, Sperren9, Sperren10, Sperren11, Sperren12, Sperren13, Sperren14, Sperren15, " +
                    "AbstvU, Farbe, BandNr, BandBr, Ged, Datum) VALUES " + "(" + GoNull[Zeile1] + ", " + GoNull[Zeile2] + ", " +
                    GoNull[Zeile3] + ", " + GoNull[Zeile4] + ", " + GoNull[Zeile5] + ", " + GoNull[Zeile6] + ", " + GoNull[Zeile7] + ", " +
                    GoNull[Zeile8] + ", " + GoNull[Zeile9] + ", " + GoNull[Zeile10] + ", " + GoNull[Zeile11] + ", " + GoNull[Zeile12] + ", " +
                    GoNull[Zeile13] + ", " + GoNull[Zeile14] + ", " + GoNull[Zeile15] + ", " + schrift + ", " +
                    GoNull[Hoehe1] + ", " + GoNull[Hoehe2] + ", " + GoNull[Hoehe3] + ", " + GoNull[Hoehe4] + ", " +
                    GoNull[Hoehe5] + ", " + GoNull[Hoehe6] + ", " + GoNull[Hoehe7] + ", " + GoNull[Hoehe8] + ", " + GoNull[Hoehe9] + ", " +
                    GoNull[Hoehe10] + ", " + GoNull[Hoehe11] + ", " + GoNull[Hoehe12] + ", " + GoNull[Hoehe13] + ", " + GoNull[Hoehe14] + ", " +
                    GoNull[Hoehe15] + ", " + GoNull[Sperren1] + ", " + GoNull[Sperren2] + ", " + GoNull[Sperren3] + ", " + GoNull[Sperren4] + ", " +
                    GoNull[Sperren5] + ", " + GoNull[Sperren6] + ", " + GoNull[Sperren7] + ", " + GoNull[Sperren8] + ", " + GoNull[Sperren9] + ", " +
                     GoNull[Sperren10] + ", " + GoNull[Sperren11] + ", " + GoNull[Sperren12] + ", " + GoNull[Sperren13] + ", " + GoNull[Sperren14] + ", " +
                     GoNull[Sperren15] + ", " + AbstandVonAussen.Value + ", " + farbe + ", " + GoNull[BandNr] + ", " + breite + ", " + Ged + ", " + Datum + ")";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            conn.Close();
            UngedruckteZeilenBerechnen();
        }

        // Leeren alle Felder
        private void EmptyTheFields()
        {
            List<TextBox> textBoxes = new List<TextBox>
            {
                Zeile1, Zeile2, Zeile3, Zeile4, Zeile5,
                Zeile6, Zeile7, Zeile8, Zeile9, Zeile10,
                Zeile11, Zeile12, Zeile13, Zeile14, Zeile15,
                Hoehe1, Hoehe2, Hoehe3, Hoehe4, Hoehe5,
                Hoehe6, Hoehe7, Hoehe8, Hoehe9, Hoehe10,
                Hoehe11, Hoehe12, Hoehe13, Hoehe14, Hoehe15,
                Sperren1, Sperren2, Sperren3, Sperren4, Sperren5,
                Sperren6, Sperren7, Sperren8, Sperren9, Sperren10,
                Sperren11, Sperren12, Sperren13, Sperren14, Sperren15,
                BandNr, Satz_Nr
            };
            foreach (TextBox t in textBoxes)
            {
                t.Text = "";
            }
            AbstandVonAussen.Value = 150;
            Breite.SelectedIndex = -1;
            FarbeEingabe.SelectedIndex = -1;
        }

        // die Anzahl der ungedruckten Zeilen berechnen
        private void UngedruckteZeilenBerechnen()
        {
            int p1 = 0, p2 = 0, p3 = 0;
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql1 = "select sum(Length(ged) - length(replace(ged, \"1\", \"\"))) from druckdatei;";
                MySqlCommand cmd1 = new MySqlCommand(sql1, conn);
                object result1 = cmd1.ExecuteScalar();
                if (result1 != null)
                {
                    p1 = Convert.ToInt32(result1);
                }
                string sql2 = "select sum(Length(ged) - length(replace(ged, \"2\", \"\"))) from druckdatei;";
                MySqlCommand cmd2 = new MySqlCommand(sql2, conn);
                object result2 = cmd2.ExecuteScalar();
                if (result2 != null)
                {
                    p2 = Convert.ToInt32(result2);
                }
                string sql3 = "select sum(Length(ged) - length(replace(ged, \"3\", \"\"))) from druckdatei;";
                MySqlCommand cmd3 = new MySqlCommand(sql3, conn);
                object result3 = cmd3.ExecuteScalar();
                if (result3 != null)
                {
                    p3 = Convert.ToInt32(result3);
                }
                InfoZeile.Text = "Ungedruckt = " + (p1 + p2 + p3).ToString() + " Zeilen     P1= " + p1.ToString() +
                 "  P2= " + p2.ToString() + "  P3= " + p3.ToString();
            }
            catch
            {
                InfoZeile.Text = "Ungedruckt = " + (p1 + p2 + p3).ToString() + " Zeilen    P1= " + p1.ToString() +
                 "  P2= " + p2.ToString() + "  P3= " + p3.ToString();
            }
            conn.Close();
        }

        // Den Druck aufrufen
        private void DruckAufrufen(string SelectDruck)
        {
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);

            try
            {
                conn.Open();
                string sql = "select * from " + AktuellDatei + " where nr = " + SelectDruck;
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    Satz_Nr.Text = rdr["nr"].ToString();
                    Zeile1.Text = rdr["Zeile1"].ToString();
                    Zeile2.Text = rdr["Zeile2"].ToString();
                    Zeile3.Text = rdr["Zeile3"].ToString();
                    Zeile4.Text = rdr["Zeile4"].ToString();
                    Zeile5.Text = rdr["Zeile5"].ToString();
                    Zeile6.Text = rdr["Zeile6"].ToString();
                    Zeile7.Text = rdr["Zeile7"].ToString();
                    Zeile8.Text = rdr["Zeile8"].ToString();
                    Zeile9.Text = rdr["Zeile9"].ToString();
                    Zeile10.Text = rdr["Zeile10"].ToString();
                    Zeile11.Text = rdr["Zeile11"].ToString();
                    Zeile12.Text = rdr["Zeile12"].ToString();
                    Zeile13.Text = rdr["Zeile13"].ToString();
                    Zeile14.Text = rdr["Zeile14"].ToString();
                    Zeile15.Text = rdr["Zeile15"].ToString();
                    Hoehe1.Text = rdr["Höhe1"].ToString();
                    Hoehe2.Text = rdr["Höhe2"].ToString();
                    Hoehe3.Text = rdr["Höhe3"].ToString();
                    Hoehe4.Text = rdr["Höhe4"].ToString();
                    Hoehe5.Text = rdr["Höhe5"].ToString();
                    Hoehe6.Text = rdr["Höhe6"].ToString();
                    Hoehe7.Text = rdr["Höhe7"].ToString();
                    Hoehe8.Text = rdr["Höhe8"].ToString();
                    Hoehe9.Text = rdr["Höhe9"].ToString();
                    Hoehe10.Text = rdr["Höhe10"].ToString();
                    Hoehe11.Text = rdr["Höhe11"].ToString();
                    Hoehe12.Text = rdr["Höhe12"].ToString();
                    Hoehe13.Text = rdr["Höhe13"].ToString();
                    Hoehe14.Text = rdr["Höhe14"].ToString();
                    Hoehe15.Text = rdr["Höhe15"].ToString();
                    Sperren1.Text = rdr["Sperren1"].ToString();
                    Sperren2.Text = rdr["Sperren2"].ToString();
                    Sperren3.Text = rdr["Sperren3"].ToString();
                    Sperren4.Text = rdr["Sperren4"].ToString();
                    Sperren5.Text = rdr["Sperren5"].ToString();
                    Sperren6.Text = rdr["Sperren6"].ToString();
                    Sperren7.Text = rdr["Sperren7"].ToString();
                    Sperren8.Text = rdr["Sperren8"].ToString();
                    Sperren9.Text = rdr["Sperren9"].ToString();
                    Sperren10.Text = rdr["Sperren10"].ToString();
                    Sperren11.Text = rdr["Sperren11"].ToString();
                    Sperren12.Text = rdr["Sperren12"].ToString();
                    Sperren13.Text = rdr["Sperren13"].ToString();
                    Sperren14.Text = rdr["Sperren14"].ToString();
                    Sperren15.Text = rdr["Sperren15"].ToString();
                    Breite.SelectedIndex = Int32.Parse(rdr["BandBr"].ToString());
                    BandNr.Text = rdr["BandNr"].ToString();
                    AbstandVonAussen.Value = Int32.Parse(rdr["AbstvU"].ToString());
                    if (Int32.Parse(rdr["Schrift"].ToString()) != -1)
                    {
                        SchriftGröße.SelectedIndex = Int32.Parse(rdr["Schrift"].ToString()) - 1;
                    }
                    else { SchriftGröße.SelectedIndex = -1; }

                    if (Int32.Parse(rdr["Farbe"].ToString()) != -1)
                    {
                        FarbeEingabe.SelectedIndex = Int32.Parse(rdr["Farbe"].ToString()) - 1;
                    }
                    else { FarbeEingabe.SelectedIndex = -1; }
                }
                rdr.Close();
            }

            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                EmptyTheFields();
            }
        }

        private void ToolStripMenuItem13_Click(object sender, EventArgs e)
        {
            AktuellDatei = "Druckdatei";
            DruckAufrufen("(select max(nr) from druckdatei);");
            AktuellDruck = Satz_Nr.Text;
            label3.Text = AktuellDatei;
        }

        private void ToolStripMenuItem12_Click(object sender, EventArgs e)
        {
            AktuellDatei = "Archiv";
            DruckAufrufen("(select max(nr) from Archiv);");
            AktuellDruck = Satz_Nr.Text;
            label3.Text = AktuellDatei;
        }

        private void VorherigerDruck_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(AktuellDruck))
            {
                toolStripMenuItem13.PerformClick();
            }
            else
            {
                try
                {
                    DruckAufrufen("(select max(nr) from " + AktuellDatei + " where nr < " + AktuellDruck + ");");
                    AktuellDruck = Satz_Nr.Text;
                }
                catch { return; }
            }
        }

        private void NächsterDruck_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(AktuellDruck))
            {
                DruckAufrufen("(select min(nr) from druckdatei);");
                AktuellDruck = Satz_Nr.Text;
            }
            else
            {
                try
                {
                    DruckAufrufen("(select min(nr) from " + AktuellDatei + " where nr > " + AktuellDruck + ");");
                    AktuellDruck = Satz_Nr.Text;
                }
                catch { return; }
            }
        }

        private void toolStripMenuItem19_Click(object sender, EventArgs e)
        {
            BandNr.Focus();
        }

        private void toolStripMenuItem20_Click(object sender, EventArgs e)
        {
            FarbeEingabe.Focus();
        }

        private void SchriftGröße_SelectedIndexChanged(object sender, EventArgs e)
        {
            Schriftgröße = SchriftGröße.SelectedIndex + 1;
            SchritGrößeÄndernFürLänge(Zeile1, LaengeZeile1);
            SchritGrößeÄndernFürLänge(Zeile2, LaengeZeile2);
            SchritGrößeÄndernFürLänge(Zeile3, LaengeZeile3);
            SchritGrößeÄndernFürLänge(Zeile4, LaengeZeile4);
            SchritGrößeÄndernFürLänge(Zeile5, LaengeZeile5);
            SchritGrößeÄndernFürLänge(Zeile6, LaengeZeile6);
            SchritGrößeÄndernFürLänge(Zeile7, LaengeZeile7);
            SchritGrößeÄndernFürLänge(Zeile8, LaengeZeile8);
            SchritGrößeÄndernFürLänge(Zeile9, LaengeZeile9);
            SchritGrößeÄndernFürLänge(Zeile10, LaengeZeile10);
            SchritGrößeÄndernFürLänge(Zeile11, LaengeZeile11);
            SchritGrößeÄndernFürLänge(Zeile12, LaengeZeile12);
            SchritGrößeÄndernFürLänge(Zeile13, LaengeZeile13);
            SchritGrößeÄndernFürLänge(Zeile14, LaengeZeile14);
            SchritGrößeÄndernFürLänge(Zeile15, LaengeZeile15);
        }

        // Hotkeys
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F6 && e.Control == true)
            {
                toolStripMenuItem10.PerformClick();
            }
            if (e.KeyCode == Keys.F6)
            {
                toolStripMenuItem11.PerformClick();
            }
            if (e.KeyCode == Keys.F7 && e.Control == true)
            {
                toolStripMenuItem12.PerformClick();
            }
            if (e.KeyCode == Keys.F7)
            {
                toolStripMenuItem13.PerformClick();
            }
            if (e.KeyCode == Keys.F8)
            {
                toolStripMenuItem14.PerformClick();
            }
            if (e.KeyCode == Keys.S && e.Control == true)
            {
                toolStripMenuItem15.PerformClick();
            }
            if (e.KeyCode == Keys.Y && e.Control == true)
            {
                toolStripMenuItem17.PerformClick();
            }
            if (e.KeyCode == Keys.F1 && e.Control == true)
            {
                toolStripMenuItem18.PerformClick();
            }
            if (e.KeyCode == Keys.F1)
            {
                toolStripMenuItem22.PerformClick();
            }
            if (e.KeyCode == Keys.F2)
            {
                toolStripMenuItem23.PerformClick();
            }
            if (e.KeyCode == Keys.F3)
            {
                toolStripMenuItem24.PerformClick();
            }
            if (e.KeyCode == Keys.F4)
            {
                toolStripMenuItem25.PerformClick();
            }
            if (e.KeyCode == Keys.F5)
            {
                toolStripMenuItem26.PerformClick();
            }
            if (e.KeyCode == Keys.F2 && e.Control == true)
            {
                toolStripMenuItem27.PerformClick();
            }
        }

        private void toolStripMenuItem14_Click(object sender, EventArgs e)
        {
            EmptyTheFields();
        }

        // Änderungen speichern
        private void DruckÄndern(string AktuellDatei)
        {
            if (CheckEveryZeileHasHöhe() == false) { return; }
            CreateNullForDatenbank();

            DateTime dt1 = DateTime.Now;
            string Datum = "\"" + dt1.ToString() + "\"";
            int breite = (Breite.Text == "") ? -1 : BreiteToDatenBank(Int32.Parse(Breite.Text));
            int farbe = (FarbeEingabe.Text == "") ? -1 : FarbeEingabe.SelectedIndex + 1;
            string Ged = "\"" + CreateGedQuerDruck() + "\"";
            string schrift = (SchriftGröße.SelectedIndex + 1).ToString();
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();
                string sql = "update " + AktuellDatei + " set Zeile1 = " + GoNull[Zeile1] +
                    ", Zeile2 = " + GoNull[Zeile2] +
                    ", Zeile3 = " + GoNull[Zeile3] +
                    ", Zeile4 = " + GoNull[Zeile4] +
                    ", Zeile5 = " + GoNull[Zeile5] +
                    ", Zeile6 = " + GoNull[Zeile6] +
                    ", Zeile7 = " + GoNull[Zeile7] +
                    ", Zeile8 = " + GoNull[Zeile8] +
                    ", Zeile9 = " + GoNull[Zeile9] +
                    ", Zeile10 = " + GoNull[Zeile10] +
                    ", Zeile11 = " + GoNull[Zeile11] +
                    ", Zeile12 = " + GoNull[Zeile12] +
                    ", Zeile13 = " + GoNull[Zeile13] +
                    ", Zeile14 = " + GoNull[Zeile14] +
                    ", Zeile15 = " + GoNull[Zeile15] +
                    ", Höhe1 = " + GoNull[Hoehe1] +
                    ", Höhe2 = " + GoNull[Hoehe2] +
                    ", Höhe3 = " + GoNull[Hoehe3] +
                    ", Höhe4 = " + GoNull[Hoehe4] +
                    ", Höhe5 = " + GoNull[Hoehe5] +
                    ", Höhe6 = " + GoNull[Hoehe6] +
                    ", Höhe7 = " + GoNull[Hoehe7] +
                    ", Höhe8 = " + GoNull[Hoehe8] +
                    ", Höhe9 = " + GoNull[Hoehe9] +
                    ", Höhe10 = " + GoNull[Hoehe10] +
                    ", Höhe11 = " + GoNull[Hoehe11] +
                    ", Höhe12 = " + GoNull[Hoehe12] +
                    ", Höhe13 = " + GoNull[Hoehe13] +
                    ", Höhe14 = " + GoNull[Hoehe14] +
                    ", Höhe15 = " + GoNull[Hoehe15] +
                    ", Sperren1 = " + GoNull[Sperren1] +
                    ", Sperren2 = " + GoNull[Sperren2] +
                    ", Sperren3 = " + GoNull[Sperren3] +
                    ", Sperren4 = " + GoNull[Sperren4] +
                    ", Sperren5 = " + GoNull[Sperren5] +
                    ", Sperren6 = " + GoNull[Sperren6] +
                    ", Sperren7 = " + GoNull[Sperren7] +
                    ", Sperren8 = " + GoNull[Sperren8] +
                    ", Sperren9 = " + GoNull[Sperren9] +
                    ", Sperren10 = " + GoNull[Sperren10] +
                    ", Sperren11 = " + GoNull[Sperren11] +
                    ", Sperren12 = " + GoNull[Sperren12] +
                    ", Sperren13 = " + GoNull[Sperren13] +
                    ", Sperren14 = " + GoNull[Sperren14] +
                    ", Sperren15 = " + GoNull[Sperren15] +
                    ", BandNr = " + GoNull[BandNr] +
                    ", schrift = " + schrift +
                    ", Datum = " + Datum +
                    ", Ged = " + Ged +
                    ", Farbe = " + farbe +
                    ", BandBr = " + breite +
                    ", AbstvU = " + AbstandVonAussen.Value +
                    " where nr = " + Satz_Nr.Text + ";";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Änderungen wurden in " + AktuellDatei + " gespeichert");
            }
            catch
            {
                MessageBox.Show("Kein Druck zum Ändern!");
            }
            conn.Close();
        }

        private void toolStripMenuItem16_Click(object sender, EventArgs e)
        {
            if (Bearbeiten_Mode)
            {
                DialogResult dr;
                dr = MessageBox.Show("Änderungen speichern?", "", MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    DruckÄndern(AktuellDatei);
                }
                else if (dr == DialogResult.No)
                {
                    return;
                }
            }
        }

        private void toolStripMenuItem17_Click(object sender, EventArgs e)
        {
            DialogResult dr;
            dr = MessageBox.Show("den Druck wirklich löschen?", "", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                try
                {
                    if (AktuellDatei == "Archiv")
                    {
                        DruckLöschen("Archiv");
                    }
                    else
                    {
                        DruckLöschen("Druckdatei");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else if (dr == DialogResult.No)
            {
                return;
            }
        }

        // Die aktuellen Druck von Druckdatei oder Archiv löschen
        private void DruckLöschen(string AktuellDatei)
        {
            string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
            MySqlConnection conn = new MySqlConnection(connStr);

            try
            {
                conn.Open();
                string sql = "delete from " + AktuellDatei + " where nr = " + Satz_Nr.Text + ";";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("den Druck wurde von " + AktuellDatei + " gelöscht");
                EmptyTheFields();
            }
            catch
            {
                MessageBox.Show("Kein Druck zum Löschen!");
            }
            conn.Close();
            UngedruckteZeilenBerechnen();
        }

        private void toolStripMenuItem18_Click(object sender, EventArgs e)
        {
            Sperren1.Focus();
        }

        // Zum Druckmode wechslen
        private void DruckMode()
        {
            this.BackColor = Color.FromArgb(38, 87, 166);
            StopButton.Visible = true;
            AutoSuchen.Visible = true;
            AutoSuchen.Text = "Autom. Suchen";
            SchriftWechseln.Visible = true;
            ReferenzFahrt.Visible = true;
            FensterWechseln.Text = "Text fenster";
            TextSpeichernOrDrucken.Text = "Drucken";
            panel1.Enabled = false;
            FarbeEingabe.Enabled = false;
            Breite.Enabled = false;
            SchriftGröße.Enabled = false;
            AbstandVonAussen.Enabled = false;
            AutomatZeilenAbstand.Enabled = false;
            DruckCheckBox.Enabled = false;
            ArchivCheckBox.Enabled = false;
            Vorschläge.Enabled = false;
            NächsterDruck.Enabled = false;
            VorherigerDruck.Enabled = false;
            AutoAbstand.Enabled = false;
            Druckzeit.Enabled = false;
            DruckStaerke.Enabled = false;
            List<TextBox> textBoxes = new List<TextBox>
            {
                Zeile1, Zeile2, Zeile3, Zeile4, Zeile5,
                Zeile6, Zeile7, Zeile8, Zeile9, Zeile10,
                Zeile11, Zeile12, Zeile13, Zeile14, Zeile15,
                Hoehe1, Hoehe2, Hoehe3, Hoehe4, Hoehe5,
                Hoehe6, Hoehe7, Hoehe8, Hoehe9, Hoehe10,
                Hoehe11, Hoehe12, Hoehe13, Hoehe14, Hoehe15,
                Sperren1, Sperren2, Sperren3, Sperren4, Sperren5,
                Sperren6, Sperren7, Sperren8, Sperren9, Sperren10,
                Sperren11, Sperren12, Sperren13, Sperren14, Sperren15,
                BandNr
            };
            foreach (TextBox t in textBoxes)
            {
                t.Enabled = false;
            }
        }

        // Zum Druckmode wechslen
        private void BearbeitenMode()
        {
            this.BackColor = SystemColors.Control;
            StopButton.Visible = false;
            AutoSuchen.Visible = false;
            AutoSuchen.Text = "Titel ändern";
            SchriftWechseln.Visible = false;
            ReferenzFahrt.Visible = false;
            FensterWechseln.Text = "Druckfenster";
            TextSpeichernOrDrucken.Text = "Text speichern";
            panel1.Enabled = true;
            FarbeEingabe.Enabled = true;
            Breite.Enabled = true;
            SchriftGröße.Enabled = true;
            AbstandVonAussen.Enabled = true;
            AutomatZeilenAbstand.Enabled = true;
            DruckCheckBox.Enabled = true;
            ArchivCheckBox.Enabled = true;
            Vorschläge.Enabled = true;
            NächsterDruck.Enabled = true;
            VorherigerDruck.Enabled = true;
            AutoAbstand.Enabled = true;
            Druckzeit.Enabled = true;
            DruckStaerke.Enabled = true;
            List<TextBox> textBoxes = new List<TextBox>
            {
                Zeile1, Zeile2, Zeile3, Zeile4, Zeile5,
                Zeile6, Zeile7, Zeile8, Zeile9, Zeile10,
                Zeile11, Zeile12, Zeile13, Zeile14, Zeile15,
                Hoehe1, Hoehe2, Hoehe3, Hoehe4, Hoehe5,
                Hoehe6, Hoehe7, Hoehe8, Hoehe9, Hoehe10,
                Hoehe11, Hoehe12, Hoehe13, Hoehe14, Hoehe15,
                Sperren1, Sperren2, Sperren3, Sperren4, Sperren5,
                Sperren6, Sperren7, Sperren8, Sperren9, Sperren10,
                Sperren11, Sperren12, Sperren13, Sperren14, Sperren15,
                BandNr
            };
            foreach (TextBox t in textBoxes)
            {
                t.Enabled = true;
            }
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            DruckSuchen frm2 = new DruckSuchen("Druckdatei");
            frm2.ShowDialog();
            AktuellDruck = stzNr;
            AktuellDatei = "Druckdatei";
            Satz_Nr.Text = stzNr;
            Zeile1.Text = z1;
            Zeile2.Text = z2;
            Zeile3.Text = z3;
            Zeile4.Text = z4;
            Zeile5.Text = z5;
            Zeile6.Text = z6;
            Zeile7.Text = z7;
            Zeile8.Text = z8;
            Zeile9.Text = z9;
            Zeile10.Text = z10;
            Zeile11.Text = z11;
            Zeile12.Text = z12;
            Zeile13.Text = z13;
            Zeile14.Text = z14;
            Zeile15.Text = z15;
            Hoehe1.Text = h1;
            Hoehe2.Text = h2;
            Hoehe3.Text = h3;
            Hoehe4.Text = h4;
            Hoehe5.Text = h5;
            Hoehe6.Text = h6;
            Hoehe7.Text = h7;
            Hoehe8.Text = h8;
            Hoehe9.Text = h9;
            Hoehe10.Text = h10;
            Hoehe11.Text = h11;
            Hoehe12.Text = h12;
            Hoehe13.Text = h13;
            Hoehe14.Text = h14;
            Hoehe15.Text = h15;
            Sperren1.Text = s1;
            Sperren2.Text = s2;
            Sperren3.Text = s3;
            Sperren4.Text = s4;
            Sperren5.Text = s5;
            Sperren6.Text = s6;
            Sperren7.Text = s7;
            Sperren8.Text = s8;
            Sperren9.Text = s9;
            Sperren10.Text = s10;
            Sperren11.Text = s11;
            Sperren12.Text = s12;
            Sperren13.Text = s13;
            Sperren14.Text = s14;
            Sperren15.Text = s15;
            BandNr.Text = BdNr;
            try { SchriftGröße.SelectedIndex = Int32.Parse(srft) - 1; } catch { SchriftGröße.SelectedIndex = -1; }
            try { Breite.SelectedIndex = Int32.Parse(brte); } catch { Breite.SelectedIndex = -1; }
            try { AbstandVonAussen.Value = Int32.Parse(AVA); } catch { AbstandVonAussen.Value = 100; }
            try { FarbeEingabe.SelectedIndex = Int32.Parse(frb) - 1; } catch { FarbeEingabe.SelectedIndex = -1; }
        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            DruckSuchen frm2 = new DruckSuchen("Archiv");
            frm2.ShowDialog();
            AktuellDatei = "Archiv";
            Satz_Nr.Text = stzNr;
            AktuellDruck = stzNr;
            Zeile1.Text = z1;
            Zeile2.Text = z2;
            Zeile3.Text = z3;
            Zeile4.Text = z4;
            Zeile5.Text = z5;
            Zeile6.Text = z6;
            Zeile7.Text = z7;
            Zeile8.Text = z8;
            Zeile9.Text = z9;
            Zeile10.Text = z10;
            Zeile11.Text = z11;
            Zeile12.Text = z12;
            Zeile13.Text = z13;
            Zeile14.Text = z14;
            Zeile15.Text = z15;
            Hoehe1.Text = h1;
            Hoehe2.Text = h2;
            Hoehe3.Text = h3;
            Hoehe4.Text = h4;
            Hoehe5.Text = h5;
            Hoehe6.Text = h6;
            Hoehe7.Text = h7;
            Hoehe8.Text = h8;
            Hoehe9.Text = h9;
            Hoehe10.Text = h10;
            Hoehe11.Text = h11;
            Hoehe12.Text = h12;
            Hoehe13.Text = h13;
            Hoehe14.Text = h14;
            Hoehe15.Text = h15;
            Sperren1.Text = s1;
            Sperren2.Text = s2;
            Sperren3.Text = s3;
            Sperren4.Text = s4;
            Sperren5.Text = s5;
            Sperren6.Text = s6;
            Sperren7.Text = s7;
            Sperren8.Text = s8;
            Sperren9.Text = s9;
            Sperren10.Text = s10;
            Sperren11.Text = s11;
            Sperren12.Text = s12;
            Sperren13.Text = s13;
            Sperren14.Text = s14;
            Sperren15.Text = s15;
            BandNr.Text = BdNr;
            try { SchriftGröße.SelectedIndex = Int32.Parse(srft) - 1; } catch { SchriftGröße.SelectedIndex = -1; }
            try { Breite.SelectedIndex = Int32.Parse(brte); } catch { Breite.SelectedIndex = -1; }
            try { AbstandVonAussen.Value = Int32.Parse(AVA); } catch { AbstandVonAussen.Value = 100; }
            try { FarbeEingabe.SelectedIndex = Int32.Parse(frb) - 1; } catch { FarbeEingabe.SelectedIndex = -1; }
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            DialogResult dr;
            dr = MessageBox.Show("Den Druckdatei wirklich leeren?", "", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                string connStr = "server=localhost;user=root;database=movedb;port=3306;password=6540";
                MySqlConnection conn = new MySqlConnection(connStr);

                try
                {
                    conn.Open();
                    string sql = "CREATE DATABASE IF NOT EXISTS movedb;USE movedb;DROP TABLE IF EXISTS Druckdatei;" +
                        "CREATE TABLE Druckdatei(nr INTEGER AUTO_INCREMENT,Zeile1 VARCHAR(100),Zeile2 VARCHAR(100)," +
                        "Zeile3 VARCHAR(100),Zeile4 VARCHAR(50),Zeile5 VARCHAR(50),Zeile6 VARCHAR(50),Zeile7 VARCHAR(50)," +
                        "Zeile8 VARCHAR(50),Zeile9 VARCHAR(50),Zeile10 VARCHAR(50),Zeile11 VARCHAR(50),Zeile12 VARCHAR(50)," +
                        "Zeile13 VARCHAR(50),Zeile14 VARCHAR(50), Zeile15 VARCHAR(50),Schrift VARCHAR(2), Höhe1 VARCHAR(5)," +
                        "Höhe2 VARCHAR(5), Höhe3 VARCHAR(5),Höhe4 VARCHAR(5),Höhe5 VARCHAR(5),Höhe6 VARCHAR(5),Höhe7 VARCHAR(5)," +
                        "Höhe8 VARCHAR(5),Höhe9 VARCHAR(5),Höhe10 VARCHAR(5),Höhe11 VARCHAR(5), Höhe12 VARCHAR(5), Höhe13 VARCHAR(5)," +
                        "Höhe14 VARCHAR(5),Höhe15 VARCHAR(5),Sperren1 VARCHAR(3),Sperren2 VARCHAR(3), Sperren3 VARCHAR(3)," +
                        "Sperren4 VARCHAR(3),Sperren5 VARCHAR(3), Sperren6 VARCHAR(3), Sperren7 VARCHAR(3), Sperren8 VARCHAR(3)," +
                        "Sperren9 VARCHAR(3), Sperren10 VARCHAR(3),Sperren11 VARCHAR(3), Sperren12 VARCHAR(3), Sperren13 VARCHAR(3)," +
                        "Sperren14 VARCHAR(3), Sperren15 VARCHAR(3), AbstvU VARCHAR(4),Farbe VARCHAR(10),BandNr VARCHAR(5)," +
                        " BandBr VARCHAR(5), Ged VARCHAR(20), Datum VARCHAR(50),INDEX(nr) ) ENGINE = myisam DEFAULT CHARSET = utf8;" +
                        "SET autocommit = 1;";

                    MySqlCommand cmd = new MySqlCommand(sql, conn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Druckdatei wurde geleert!");
                    EmptyTheFields();
                    UngedruckteZeilenBerechnen();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                conn.Close();
            }
            else
            {
                return;
            }
        }

        // zweite Seite bestimmen
        private void Zeile1_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 0;
            Seite2Waehlen();
        }

        private void Zeile2_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 2;
            Seite2Waehlen();
        }

        private void Zeile3_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 3;
            Seite2Waehlen();
        }

        private void Zeile4_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 4;
            Seite2Waehlen();
        }

        private void Zeile5_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 5;
            Seite2Waehlen();
        }

        private void Zeile6_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 6;
            Seite2Waehlen();
        }

        private void Zeile7_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 7;
            Seite2Waehlen();
        }

        private void Zeile8_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 8;
            Seite2Waehlen();
        }

        private void Zeile9_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 9;
            Seite2Waehlen();
        }

        private void Zeile10_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 10;
            Seite2Waehlen();
        }

        private void Zeile11_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 11;
            Seite2Waehlen();
        }

        private void Zeile12_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 12;
            Seite2Waehlen();
        }

        private void Zeile13_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 13;
            Seite2Waehlen();
        }

        private void Zeile14_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 14;
            Seite2Waehlen();
        }

        private void Zeile15_DoubleClick(object sender, EventArgs e)
        {
            Seite2 = 15;
            Seite2Waehlen();
        }

        private void Seite2Waehlen()
        {
            List<TextBox> Hoehen = new List<TextBox>
            {
                Hoehe1, Hoehe2, Hoehe3, Hoehe4, Hoehe5,
                Hoehe6, Hoehe7, Hoehe8, Hoehe9, Hoehe10,
                Hoehe11, Hoehe12, Hoehe13, Hoehe14, Hoehe15
            };
            if (Seite2 == 0)
            {
                for (int i = Seite2; i < 15; i++)
                {
                    Hoehen[i].BackColor = Color.White;
                }
                return;
            }
            for (int i = Seite2; i < 16; i++)
            {
                Hoehen[i - 1].BackColor = Color.LightGray;
            }
            for (int i = Seite2; i > 1; i--)
            {
                Hoehen[i - 2].BackColor = Color.White;
            }
        }

        private void AutoAbstand_Click(object sender, EventArgs e)
        {
            int Abstand = Decimal.ToInt32(AbstandVonAussen.Value);
            int Automat = Decimal.ToInt32(AutomatZeilenAbstand.Value);
            int st1 = 0, st2 = 0;
            List<TextBox> Hoehen = new List<TextBox>
            {
                InfoZeile,Hoehe1, Hoehe2, Hoehe3, Hoehe4, Hoehe5,
                Hoehe6, Hoehe7, Hoehe8, Hoehe9, Hoehe10,
                Hoehe11, Hoehe12, Hoehe13, Hoehe14, Hoehe15
            };

            List<TextBox> Zeilen = new List<TextBox>
            {
                InfoZeile,Zeile1, Zeile2, Zeile3, Zeile4, Zeile5,
                Zeile6, Zeile7, Zeile8, Zeile9, Zeile10,
                Zeile11, Zeile12, Zeile13, Zeile14, Zeile15
            };
            if (Seite2 == 0)
            {   
                for(int i = 15; i >= 1; i--)
                {
                    if (Zeilen[i].Text == "") continue;
                    else
                    {
                        Hoehen[i].Text = Abstand.ToString();
                        Abstand += Automat;
                    }
                }
            }
            else
            {
                for (int i = Seite2; i <= 15; i++)
                {
                    if (Zeilen[i].Text == "") continue;
                    else { st2 += 1; }
                }
                for (int i = Seite2 - 1; i > 0; i--)
                {
                    if (Zeilen[i].Text == "") continue;
                    else { st1 += 1; }
                }
                if (st1 >= st2 && st1 != 0 && st2 !=0)
                {
                    for (int i = Seite2 - 1; i >= 1; i--)
                    {
                        if (Zeilen[i].Text == "") continue;
                        else
                        {
                            Hoehen[i].Text = Abstand.ToString();
                            Abstand += Automat;
                        }
                    }
                    Abstand -= Automat;
                    decimal Mitte = (AbstandVonAussen.Value + Abstand) / 2;
                    decimal x = 0;
                    for (int i = 1; i < st2; i++)
                    {
                        x += i;
                    }
                    decimal ZH = Mitte - (x / st2) * AutomatZeilenAbstand.Value;                    
                    for (int i = 15; i >= Seite2; i--)
                    {
                        if (Zeilen[i].Text == "") continue;
                        else
                        {
                            Hoehen[i].Text = (decimal.ToInt32(ZH)).ToString();
                            ZH += AutomatZeilenAbstand.Value;
                        }
                    }
                }
                else if (st1 < st2 && st1 != 0 && st2 != 0)
                {
                    for (int i = 15; i >= Seite2; i--)
                    {
                        if (Zeilen[i].Text == "") continue;
                        else
                        {
                            Hoehen[i].Text = Abstand.ToString();
                            Abstand += Automat;
                        }
                    }
                    Abstand -= Automat;
                    decimal Mitte = (AbstandVonAussen.Value + Abstand) / 2;
                    decimal x = 0;
                    for (int i = 1; i < st1; i++)
                    {
                        x += i;
                    }
                    decimal ZH = Mitte - (x / st1) * AutomatZeilenAbstand.Value;
                    for (int i = Seite2 - 1; i >= 1; i--)
                    {
                        if (Zeilen[i].Text == "") continue;
                        else
                        {
                            Hoehen[i].Text = (decimal.ToInt32(ZH)).ToString();
                            ZH += AutomatZeilenAbstand.Value;
                        }
                    }
                }
            }
            HoehenLeeren();
        }
        
        // Die Höhe der leeren Zeilen leeren wenn man auf Autom. Zeilenabstand kliclt
        private void HoehenLeeren()
        {
            List<TextBox> Zeilen = new List<TextBox>
            {
                Zeile1, Zeile2, Zeile3, Zeile4, Zeile5,
                Zeile6, Zeile7, Zeile8, Zeile9, Zeile10,
                Zeile11, Zeile12, Zeile13, Zeile14, Zeile15
            };
            List<TextBox> Hoehen = new List<TextBox>
            {
                Hoehe1, Hoehe2, Hoehe3, Hoehe4, Hoehe5,
                Hoehe6, Hoehe7, Hoehe8, Hoehe9, Hoehe10,
                Hoehe11, Hoehe12, Hoehe13, Hoehe14, Hoehe15
            };
            for (int i = 0; i <= 14; i++)
            {
                if (Zeilen[i].Text == "") Hoehen[i].Text = "";
            }
        }
    }
}
