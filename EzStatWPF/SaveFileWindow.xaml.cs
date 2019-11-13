using mshtml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
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
    /// Логика взаимодействия для SaveFileWindow.xaml
    /// </summary>
    public partial class SaveFileWindow : Window
    {
        public string zfullname;
        public Dictionary<string, string> data = new Dictionary<string, string>();
        public string zvitid;
        public string htmldoc = "";
        public HTMLDocument hdoc = null;

        public SaveFileWindow()
        {
            InitializeComponent();
            
        }
        //save file
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            using (OleDbConnection connection = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = EsDb.mdb;"))
            {
                bool canSave = true;

                OleDbCommand command2 = new OleDbCommand("Select saveName From D_saves");
                command2.Connection = connection;               
                connection.Open();
                OleDbDataReader reader = command2.ExecuteReader();
                while (reader.Read())
                {
                    if (saveName.Text.Equals(reader[0].ToString())) { canSave = false; break; }
                }
                connection.Close();

                if (canSave)
                {
                    string s = "";
                    List<string> etables = new List<string>();
                    IHTMLElement el = null;
                    for (int i = 0; i < 99; ++i)
                    {
                        if (hdoc.getElementById("edittable" + i) != null)
                         el = hdoc.getElementById("edittable" + i);

                        if (el != null)
                        {
                            etables.Add(el.outerHTML); el = null;
                        }                          
                    }

                    string htmldata = "";
                    OleDbCommand gcmd = new OleDbCommand
                    {
                        CommandType = CommandType.Text,
                        CommandText = "Select zvitView From dovZvitForm Where zvitShortName='"+zvitid+"'",
                        Connection = connection
                    };
                    try
                    {
                        connection.Open();
                        reader = gcmd.ExecuteReader();
                        while (reader.Read())
                        {
                            File.WriteAllBytes("kss.html", (byte[])reader[0]);
                        }
                            connection.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                    htmldata = File.ReadAllText("kss.html");
                    int k = 0;
                    foreach (string ss in etables)
                    {
                        string hd = htmldata;                    
                        hd = hd.Remove(hd.IndexOf("<!--START" + k + "-->")+13,(hd.IndexOf("<!--END"+k+"-->")-1)-(hd.IndexOf("<!--START" + k + "-->") + 13));
                        hd = hd.Insert(hd.IndexOf("<!--START" + k + "-->") + 13, ss);
                            ++k; htmldata = hd;
                    } // end edit html


                    using (StreamWriter writetext = new StreamWriter("save2.html"))
                    {
                      writetext.WriteLine(htmldata);                        
                    }//

                    List<string> keys = new List<string>(data.Keys);
                    foreach (string key in keys)
                    {
                        if (data[key] != null)
                            data[key] = data[key].Replace("\n", " $$a ");
                    }
                    using (StreamWriter writetext = new StreamWriter("save1.txt"))
                    {
                        foreach (KeyValuePair<string, string> kp in data)
                        {
                            writetext.WriteLine("id " + kp.Key + " val " + kp.Value);
                        }
                    }

                    byte[] b1 = File.ReadAllBytes("save1.txt");
                    byte[] b2 = File.ReadAllBytes("save2.html");

                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "insert into D_saves ([saveName],[zvitId],[zDescr],[zDate],[zTextData],[zHtmlData]) values (?,?,?,?,?,?)";
                    cmd.Parameters.AddWithValue("@saveName", saveName.Text);
                    cmd.Parameters.AddWithValue("@zvitId", zvitid.Substring(0, 2));
                    if (saveDescr.Text != "")
                        cmd.Parameters.AddWithValue("@zDecr", saveDescr.Text);
                    else
                        cmd.Parameters.AddWithValue("@zDecr", "Без опису");
                    cmd.Parameters.AddWithValue("@zDate", DateTime.Today.ToString("dd.MM.yyyy"));
                    cmd.Parameters.AddWithValue("@zTextData", b1);
                    cmd.Parameters.AddWithValue("@zHtmlData", b2);
                    cmd.Connection = connection;
                    try
                    {
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Звіт успішно збережено", "Операція збереження");
                        connection.Close();
                        this.Close(); // close form after sucessfull saving
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Невдача при збережені" + ex.Message, "Операція збереження ");
                    }
                    //clear temp files
                    File.Delete("kss.html"); 
                    File.Delete("save1.txt");
                    File.Delete("save2.html");
                }   //end if cansave      
                else
                MessageBox.Show("Файл з таким ім'ям уде збережено");
            }//end using            
        }//end fun

    }
}

