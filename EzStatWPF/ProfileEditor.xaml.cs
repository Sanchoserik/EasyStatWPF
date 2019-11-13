using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace EzStatWPF
{
    public partial class ProfileEditor : Window
    {
       public CProfile pr;

        public ProfileEditor()
        {
            InitializeComponent();
           
            string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand cmd = new OleDbCommand("Select * from KOATUUX", con);
            OleDbCommand cmd2 = new OleDbCommand("Select * from D_kved", con);
            con.Open();
            OleDbDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {              
                KOATUU.Items.Add(reader[0].ToString().Substring(0, 2) + "  " + reader[0].ToString().Substring(2, 3) + " " + reader[2].ToString());               
            }
            OleDbDataReader reader2 = cmd2.ExecuteReader();
            while (reader2.Read())
            {
                KVED.Items.Add(reader2[0].ToString() + "  " + reader2[1].ToString());
            }
            con.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {           

            if (pr.FIZ_O.Equals("1")) Fiz_y.IsChecked = true;
            else Fiz_n.IsChecked = true;
            Firm_name.Text = pr.FIRM_NAME;
            Firm_adr.Text = pr.FIRM_ADR;
            Firm_adr_fiz.Text = pr.FIRM_ADR_FIZ;
            Firm_email.Text = pr.VIK_EMAIL;
            Firm_faxorg.Text = pr.FIRM_FAXORG;
            Vik.Text = pr.VIK;
            Vik_ruk.Text = pr.VIK_RUK;
            Vik_tel.Text = pr.VIK_TEL;
            Edrpou.Text = pr.EDRPOU;

            C_reg.Text = pr.C_REG;
            C_raj.Text = pr.C_RAJ;
            Firm_spato.Text = pr.FIRM_SPATO;
            Firm_kved.Text = pr.FIRM_KVED;

        }

        private void Window_Closed(object sender, EventArgs e)
        {
            // Application.Current.Shutdown();
            MainWindow mw = new MainWindow();
            mw.Show();
        }

        //searchKOATUU
        private void SearchKOATUU_TextChanged(object sender, TextChangedEventArgs e)
        {
            string command;
            if (!SearchKOATUU.Text.Equals(""))
            {
                command = "Select * from KOATUUX WHERE NU LIKE '" + SearchKOATUU.Text + "%'";
            }
            else
            {
                command = "Select * from KOATUUX ";
            }
            KOATUU.Items.Clear();
            string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";           
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand cmd = new OleDbCommand(command, con);
            con.Open();
            OleDbDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                KOATUU.Items.Add(reader[0].ToString().Substring(0, 2) + "  " + reader[0].ToString().Substring(2, 3) + " " + reader[2].ToString());
            }
            con.Close();
        }
        //search KVED
        private void SearchKVED_TextChanged(object sender, TextChangedEventArgs e)
        {
            string command;
            if (!SearchKVED.Text.Equals(""))
            {
                command = "Select * from D_kved WHERE NU LIKE '" + SearchKVED.Text + "%'";
            }
            else
            {
                command = "Select * from D_kved ";
            }
            KVED.Items.Clear();
            string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";            
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand cmd = new OleDbCommand(command, con);
            con.Open();
            OleDbDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                KVED.Items.Add(reader[0].ToString() + "  " + reader[1].ToString());
            }
            con.Close();
        }

        //select KOATUU
        private void KOATUU_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (KOATUU.SelectedIndex != -1)
            {
                C_reg.Text = KOATUU.SelectedItem.ToString().Substring(0, 2);
                C_raj.Text = KOATUU.SelectedItem.ToString().Substring(4, 3);
                string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";
                OleDbConnection con = new OleDbConnection(cs);
                string x = KOATUU.SelectedItem.ToString().Substring(8, KOATUU.SelectedItem.ToString().Length - 8);
                OleDbCommand cmd = new OleDbCommand("Select * from KOATUUX WHERE NU='" + x + "'", con);
                con.Open();
                OleDbDataReader reader = cmd.ExecuteReader();
                reader.Read();
                Firm_spato.Text = reader[0].ToString();
                con.Close();
            }
        }
        //
        private void KVED_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(KVED.SelectedIndex !=-1)
            Firm_kved.Text = KVED.SelectedItem.ToString().Substring(0, KVED.SelectedItem.ToString().IndexOf(" "));
        }

        //return with no changes
        private void ReturnWithNoChanges_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mw = new MainWindow();
            mw.Show();
            this.Close();
        }

        //return with changes
        private void ApplyChanges_Click(object sender, RoutedEventArgs e)
        {
            // validation rules
            bool pvalid = true;
            if (Edrpou.Text == "" || Firm_name.Text == "" || Firm_adr.Text == "" || Firm_adr_fiz.Text == ""
                || Vik_ruk.Text == "" || Vik.Text == "" || Vik_tel.Text == "" || C_reg.Text == "" || C_raj.Text == ""
                || Firm_spato.Text == "" || Firm_kved.Text == "" || Firm_faxorg.Text == "")
            { MessageBox.Show("Усі поля мають бути заповнені"); pvalid = false; }
            else if (Edrpou.Text.Length != 8)
            { MessageBox.Show("Поле ЄДРПОУ повинно мати довжину 8 символів"); pvalid = false; }
            else if (C_reg.Text.Length != 2)
            { MessageBox.Show("Поле код області повинно мати довжину 2 символів"); pvalid = false; }
            else if (C_raj.Text.Length != 3)
            { MessageBox.Show("Поле код району повинно мати довжину 3 символів"); pvalid = false; }
            else if (Firm_spato.Text.Length != 10)
            { MessageBox.Show("Поле код території повинно мати довжину 10 символів"); pvalid = false; }

            if(pvalid)
            using (OleDbConnection connection = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = EsDb.mdb;"))
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Profiles SET EDRPOU = ?, FIRM_NAME = ?, FIRM_ADR = ?, FIRM_ADR_FIZ =?, VIK_RUK = ?, VIK = ?, VIK_TEL = ?, VIK_EMAIL = ?, FIRM_FAXOR = ?,"+
                    "FIRM_KVED = ?, FIRM_SPATO = ?, C_REG = ?, C_RAJ = ?, FIZ_O = ? WHERE Profile ='" + pr.name + "'";                
                cmd.Parameters.AddWithValue("@EDRPOU", Convert.ToInt32(Edrpou.Text));
                cmd.Parameters.AddWithValue("@FIRM_NAME", Firm_name.Text);
                cmd.Parameters.AddWithValue("@FIRM_ADR", Firm_adr.Text);
                cmd.Parameters.AddWithValue("@FIRM_ADR_FIZ", Firm_adr_fiz.Text);
                cmd.Parameters.AddWithValue("@VIK_RUK", Vik_ruk.Text);
                cmd.Parameters.AddWithValue("@VIK", Vik.Text);
                cmd.Parameters.AddWithValue("@VIK_TEL", Vik_tel.Text);
                cmd.Parameters.AddWithValue("@VIK_EMAIL", Firm_email.Text);
                cmd.Parameters.AddWithValue("@FIRM_FAXOR", Firm_faxorg.Text);
                cmd.Parameters.AddWithValue("@FIRM_KVED", Firm_kved.Text);
                cmd.Parameters.AddWithValue("@FIRM_SPATO", Firm_spato.Text);
                cmd.Parameters.AddWithValue("@C_REG", C_reg.Text);
                cmd.Parameters.AddWithValue("@C_RAJ", C_raj.Text);
                if(Fiz_y.IsChecked == true)
                    cmd.Parameters.AddWithValue("@FIZ_O", "1");
                else
                    cmd.Parameters.AddWithValue("@FIZ_O", "0");

                cmd.Connection = connection;

                try
                {
                    connection.Open();
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Профіль успішно оновлено", "Операція оновлення");
                    connection.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

          
            this.Close();
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }
    }
}
