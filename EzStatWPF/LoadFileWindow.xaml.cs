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
using System.Windows.Shapes;

namespace EzStatWPF
{
    /// <summary>
    /// Логика взаимодействия для LoadFileWindow.xaml
    /// </summary>
    public partial class LoadFileWindow : Window
    {
        public MainForm mf;

        public LoadFileWindow()
        {
            InitializeComponent();

            string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand cmd = new OleDbCommand("Select * from D_saves", con);         
            con.Open();
            OleDbDataReader reader = cmd.ExecuteReader();
            List<string> lyears = new List<string>();
            while (reader.Read())
            {
                SaveList.Items.Add(reader[0].ToString());

                string s = reader[3].ToString().Substring(6,4);
                if (lyears.Find(x => x.Equals(s)) == null)
                    lyears.Add(reader[3].ToString().Substring(6, 4));
                    
            }
            con.Close();

            zYear.Items.Add("Всі");
            foreach (string s in lyears)
                zYear.Items.Add(s);
            zYear.SelectedIndex = 0;

            zPeriod.Items.Add("Всі");

            zPeriod.Items.Add("Всі Місяці");
            zPeriod.Items.Add("Січень");
            zPeriod.Items.Add("Лютий");
            zPeriod.Items.Add("Березень");
            zPeriod.Items.Add("Квітень");
            zPeriod.Items.Add("Травень");
            zPeriod.Items.Add("Червень");
            zPeriod.Items.Add("Липень");
            zPeriod.Items.Add("Серпень");
            zPeriod.Items.Add("Вересень");
            zPeriod.Items.Add("Жовтень");
            zPeriod.Items.Add("Листопад");
            zPeriod.Items.Add("Грудень");

            zPeriod.Items.Add("I Квартал");
            zPeriod.Items.Add("II Квартал");
            zPeriod.Items.Add("III Квартал");
            zPeriod.Items.Add("IV Квартал");

            zPeriod.Items.Add("6 Місяців(перше півріччя)");

            zPeriod.Items.Add("9 Місяців");

            zPeriod.Items.Add("Рік");

            zPeriod.SelectedIndex = 0;
        }

        private void Load_Click(object sender, RoutedEventArgs e)
        {
            string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand cmd = new OleDbCommand("Select * from D_saves WHERE saveName='" + SaveList.SelectedItem.ToString() + "'", con);
            con.Open();
            OleDbDataReader reader = cmd.ExecuteReader();
            reader.Read();
            //start reader
            string zvitid = reader[1].ToString();
            mf.loadZvit(zvitid, SaveList.SelectedItem.ToString());
            //end reader
            con.Close();
            this.Close();
        }

        //change save
        private void SaveList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SaveList.SelectedIndex != -1)
            {              
                string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";
                OleDbConnection con = new OleDbConnection(cs);               
                OleDbCommand cmd = new OleDbCommand("Select * from D_saves WHERE saveName='" + SaveList.SelectedItem.ToString() + "'", con);
                con.Open();
                OleDbDataReader reader = cmd.ExecuteReader();
                reader.Read();
                sdescr.Text = reader[2].ToString();
                sdate.Content = reader[3].ToString();
                con.Close();
            }
        }

        //delete save
        private void DeleteSave_Click(object sender, RoutedEventArgs e)
        {
            string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand cmd = new OleDbCommand("Delete from D_saves WHERE saveName='" + SaveList.SelectedItem.ToString() + "'", con);
            
            con.Open();
            cmd.ExecuteNonQuery();

            OleDbCommand cmd2 = new OleDbCommand("Select saveName from D_saves", con);
            SaveList.Items.Clear();
            SaveList.SelectedIndex = -1;
            sdate.Content = ""; sdescr.Text = "";
            OleDbDataReader reader = cmd2.ExecuteReader();
            while (reader.Read())
            {
                SaveList.Items.Add(reader[0].ToString());
            }

            con.Close(); 
        }

        //apply filter
        private void FilterSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int year = 0;
            if (zYear.SelectedIndex != 0)
                year = Convert.ToInt32(zYear.SelectedItem);
           
            //
            SaveList.Items.Clear();
            List<savecell> scells = new List<savecell>();
            List<savecell> scelltrue = new List<savecell>();

            string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand cmd = new OleDbCommand("Select * from D_saves", con);
            con.Open();
            OleDbDataReader reader = cmd.ExecuteReader();          
            while (reader.Read())
            {
                scells.Add(new savecell { zname = reader[0].ToString(), zdate = reader[3].ToString(), zid = reader[1].ToString() });            
            }
            con.Close();
   
            con.Open();
            foreach (savecell sc in scells)
            {
                OleDbCommand cmd2 = new OleDbCommand("Select zvitPeriod from dovZvitForm where zvitShortName='"+sc.zid+"'", con);
                reader = cmd2.ExecuteReader();
                while (reader.Read())
                {
                    sc.zperiod = reader[0].ToString(); }}
            con.Close();

            List<string> namesToRemove = new List<string>();
            if (year != 0)
                foreach (savecell sc in scells)
                {
                    if (Convert.ToInt32(sc.zdate.Substring(6, 4)) != year)
                        namesToRemove.Add(sc.zname);
                }
           
            if (zPeriod.SelectedIndex != 0)
                foreach (savecell sc in scells)
                {                    
                  switch (zPeriod.SelectedIndex)
                    {
                        case 1: { if (sc.zperiod != "1")namesToRemove.Add(sc.zname);  break; }
                        case 2: { if (sc.zdate.Substring(3, 2) != "01") namesToRemove.Add(sc.zname); break; }
                        case 3: { if (sc.zdate.Substring(3, 2) != "02") namesToRemove.Add(sc.zname); break; }
                        case 4: { if (sc.zdate.Substring(3, 2) != "03") namesToRemove.Add(sc.zname); break; }
                        case 5: { if (sc.zdate.Substring(3, 2) != "04") namesToRemove.Add(sc.zname); break; }
                        case 6: { if (sc.zdate.Substring(3, 2) != "05") namesToRemove.Add(sc.zname); break; }
                        case 7: { if (sc.zdate.Substring(3, 2) != "06") namesToRemove.Add(sc.zname); break; }
                        case 8: { if (sc.zdate.Substring(3, 2) != "07") namesToRemove.Add(sc.zname); break; }
                        case 9: { if (sc.zdate.Substring(3, 2) != "08") namesToRemove.Add(sc.zname); break; }
                        case 10: { if (sc.zdate.Substring(3, 2) != "09") namesToRemove.Add(sc.zname); break; }
                        case 11: { if (sc.zdate.Substring(3, 2) != "10") namesToRemove.Add(sc.zname); break; }
                        case 12: { if (sc.zdate.Substring(3, 2) != "11") namesToRemove.Add(sc.zname); break; }
                        case 13: { if (sc.zdate.Substring(3, 2) != "12") namesToRemove.Add(sc.zname); break; }
                        case 14: { if (sc.zdate.Substring(3, 2) != "01" || sc.zdate.Substring(3, 2) != "02" || sc.zdate.Substring(3, 2) != "03") namesToRemove.Add(sc.zname); break; }
                        case 15: { if (sc.zdate.Substring(3, 2) != "04" || sc.zdate.Substring(3, 2) != "05" || sc.zdate.Substring(3, 2) != "06") namesToRemove.Add(sc.zname); break; }
                        case 16: { if (sc.zdate.Substring(3, 2) != "07" || sc.zdate.Substring(3, 2) != "08" || sc.zdate.Substring(3, 2) != "09") namesToRemove.Add(sc.zname); break; }
                        case 17: { if (sc.zdate.Substring(3, 2) != "10" || sc.zdate.Substring(3, 2) != "11" || sc.zdate.Substring(3, 2) != "12") namesToRemove.Add(sc.zname); break; }
                        case 18: { if (sc.zperiod != "3") namesToRemove.Add(sc.zname); break; }
                        case 19: { if (sc.zperiod != "4") namesToRemove.Add(sc.zname); break; }
                        case 20: { if (sc.zperiod != "5") namesToRemove.Add(sc.zname); break; }
                    }//end switch
                }

            foreach (string s in namesToRemove)
                scells.RemoveAll(x => x.zname.Equals(s));

            foreach (savecell sc in scells)
            {
                SaveList.Items.Add(sc.zname);
            }

        }

        private class savecell
        {
            public string zname;
            public string zid;
            public string zdate;
            public string zperiod;
        }
    }//end class
}
