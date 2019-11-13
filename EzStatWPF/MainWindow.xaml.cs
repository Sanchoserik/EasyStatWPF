using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace EzStatWPF
{
    public partial class MainWindow : Window
    {
        CProfile profile = null;

        public MainWindow()
        {
            InitializeComponent();
        

            //get profiles
            string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";           
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand command = new OleDbCommand("Select Profile From Profiles", con);

            con.Open();
            OleDbDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                Profiles.Items.Add(reader[0].ToString());
            }
            con.Close();
        }

        // select profile
        private void Profiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            profile = new CProfile();

            string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand command = new OleDbCommand("Select * From Profiles Where Profile='"+Profiles.SelectedItem.ToString()+"'", con);

            con.Open();
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read(); // read first line 
            profile.name = reader[1].ToString(); 
            profile.EDRPOU = reader[2].ToString(); Edrpou.Content = profile.EDRPOU;
            profile.FIRM_NAME = reader[3].ToString(); Firm_name.Content = profile.FIRM_NAME;
            profile.FIRM_ADR = reader[4].ToString(); Firm_adr.Content = profile.FIRM_ADR;
            profile.FIRM_ADR_FIZ = reader[5].ToString(); Firm_adr_fiz.Content = profile.FIRM_ADR_FIZ;
            profile.VIK_RUK = reader[6].ToString(); Vik_ruk.Content = profile.VIK_RUK;
            profile.VIK = reader[7].ToString(); Vik.Content = profile.VIK;
            profile.VIK_TEL = reader[8].ToString(); Vik_tel.Content = profile.VIK_TEL;
            profile.VIK_EMAIL = reader[9].ToString(); Firm_email.Content = profile.VIK_EMAIL;
            profile.FIRM_FAXORG = reader[10].ToString(); Firm_faxorg.Content = profile.FIRM_FAXORG;
            profile.FIRM_KVED = reader[11].ToString(); Firm_kved.Content = profile.FIRM_KVED;
            profile.FIRM_SPATO = reader[12].ToString(); Firm_spato.Content = profile.FIRM_SPATO;
            profile.C_RAJ = reader[13].ToString(); C_raj.Content = profile.C_RAJ;
            profile.C_REG = reader[14].ToString(); C_reg.Content = profile.C_REG;
            profile.FIZ_O = reader[15].ToString(); Fiz_o.Content = profile.FIZ_O;

            con.Close();
        }

        //go to Profile editor 
        private void ProfileEditor_Click(object sender, RoutedEventArgs e)
        {
            if (Profiles.SelectedIndex != -1)
            {
                ProfileEditor ed = new ProfileEditor();
                ed.pr = profile;
                ed.Show();
                this.Hide();
            }
        }

        //
        private void ProfileSelect_Closed(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void ProfileEditor_Copy_Click(object sender, RoutedEventArgs e)
        {
            ProfileCreator pc = new ProfileCreator();
            pc.Show();
            this.Hide();
        }

        //go to main form
        private void SelectProfile_Click(object sender, RoutedEventArgs e)
        {
            MainForm mf = new MainForm();
            mf.profile = profile;
            mf.WindowState = WindowState.Maximized;
            mf.Show();
            this.Hide();
        }
    }
}
