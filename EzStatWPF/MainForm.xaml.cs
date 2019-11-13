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
using EzStatWPF.Zforms;
using System.Xml.Serialization;
using System.Threading;
using System.Windows.Navigation;

namespace EzStatWPF
{
    /// <summary>
    /// Логика взаимодействия для MainForm.xaml
    /// </summary>
    public partial class MainForm : Window
    {
        public CProfile profile = null; // profile data
        List<string> dataList = new List<string>(); // data storage for current tab
        public int counter = 0; // count tabs

        public MainForm()
        {
            InitializeComponent();
            this.WindowState = WindowState.Maximized;
              
        }

        //zoom
        public void Slide(object sender, EventArgs e)
        {
            Slider sd = (Slider)sender;            
            TabItem x = TabPanel.SelectedItem as TabItem;
            List<Visual> clist = VisualTreeEnumHelper.GetAllControls((Visual)x.Content);
            foreach (Visual v in clist)
            {
                if (v is WebBrowser)
                {                   
                    WebBrowser wb = (WebBrowser)v;
                    mshtml.IHTMLDocument2 doc = wb.Document as mshtml.IHTMLDocument2;
                    doc.parentWindow.execScript("document.body.style.zoom=" + sd.Value.ToString().Replace(",", ".") + ";");
                    break; // interrupt
                }
            }

        }

        //create open dialog
        private void OpenFileDialog(object sender, RoutedEventArgs e)
        {
            StartNewZvit sn = new StartNewZvit();
            sn.mf = this;
            sn.Show();
        }

        //load zvit
        public void loadZvit(string shortname, string savename)
        {
            createWebbrowserTab(shortname, savename);
            loadZvit(savename);
        }
        //start new zvit
        public void startNewZvit(string shortname, string longname)
        {
            createWebbrowserTab(shortname, longname);
            getZvit(shortname); // wb navigate to zvit
        }
        //create webBrowsertab
        public void createWebbrowserTab(string shortname, string longname)
        {            
            Grid gr = new Grid();
            //gr.Background = Brushes.Azure;

            gr.RowDefinitions.Add(new RowDefinition { Height = new GridLength(DockElement.ActualHeight * 0.8) });
            gr.RowDefinitions.Add(new RowDefinition { Height = new GridLength(DockElement.ActualHeight * 0.01) });
            gr.RowDefinitions.Add(new RowDefinition { Height = new GridLength(DockElement.ActualHeight * 0.19) });

            gr.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });
            gr.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(DockElement.ActualWidth - 40) });
            gr.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });

            gr.HorizontalAlignment = HorizontalAlignment.Stretch;
            gr.VerticalAlignment = VerticalAlignment.Stretch;

            //shadow label
            Label nameX = new Label();
            nameX.Content = shortname + " " + counter;
            counter++;
            nameX.Visibility = Visibility.Hidden;
            nameX.Width = 1; nameX.Height = 1;
            nameX.Margin = new Thickness(1, 1, 0, 0);
            Grid.SetRow(nameX, 0);
            Grid.SetColumn(nameX, 0);
            gr.Children.Add(nameX);
            dataList.Add(nameX.Content.ToString());
            //end shadow label
            
            WebBrowser wb = new WebBrowser();
            wb.HorizontalAlignment = HorizontalAlignment.Stretch;
            wb.VerticalAlignment = VerticalAlignment.Stretch;
     
            Grid.SetRow(wb, 0);
            Grid.SetColumn(wb, 1);
            gr.Children.Add(wb);

            Slider sd = new Slider();
            sd.HorizontalAlignment = HorizontalAlignment.Right;
            sd.VerticalAlignment = VerticalAlignment.Stretch;
            sd.Height = DockElement.ActualHeight * 0.19; sd.Width = 100;
            sd.Minimum = 0; sd.Maximum = 5; sd.Value = 2;
            sd.SmallChange = 0.1;
            sd.LargeChange = 1;
            //sd.Background = Brushes.ForestGreen;
            
            sd.AddHandler(Slider.ValueChangedEvent, new RoutedEventHandler(Slide));
            Grid.SetRow(sd, 2);
            Grid.SetColumn(sd, 1);
            gr.Children.Add(sd);

            TabPanel.Items.Add(new TabItem
            {
                Header = new TextBlock { Text = longname }, // установка заголовка вкладки
                Content = gr// установка содержимого вкладки
            });

            TabPanel.SelectedItem = TabPanel.Items[TabPanel.Items.Count - 1]; // automaticaly select         
        }//end fun

        //get zvit by short name ID
        private void getZvit(string id)
        {
            using (OleDbConnection connection = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = EsDb.mdb;"))
            {
                OleDbCommand command = new OleDbCommand("Select zvitView From dovZvitForm Where zvitShortName='" + id + "'");
                command.Connection = connection;
                try
                {
                    connection.Open(); 
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {                       
                        File.WriteAllBytes("f.html", (byte[])reader[0]);
                        string fullpath = System.IO.Path.GetFullPath("f.html");

                        // get webbrowser
                        TabItem x = this.TabPanel.SelectedItem as TabItem;
                        List<Visual> clist = VisualTreeEnumHelper.GetAllControls((Visual)x.Content);
                        foreach (Visual v in clist)
                            if (v is WebBrowser)
                            {
                                WebBrowser wb = v as WebBrowser;
                                wb.Navigate(fullpath);
                                break;
                            }
             
                    }
                    reader.Close();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }                
            }
        }//end get zvit
        //load zvit ex
        private void loadZvit(string id)
        {
            TabItem x = this.TabPanel.SelectedItem as TabItem;
            List<Visual> clist = VisualTreeEnumHelper.GetAllControls((Visual)x.Content);
          
            using (OleDbConnection connection = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = EsDb.mdb;"))
            {
                OleDbCommand command = new OleDbCommand("Select * From D_saves Where saveName='" + id + "'");
                command.Connection = connection;
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        File.WriteAllBytes("s.txt", (byte[])reader[4]);
                        File.WriteAllBytes("fl.html", (byte[])reader[5]);
                        string fullpath = System.IO.Path.GetFullPath("fl.html");

                        // get webbrowser
                        x = this.TabPanel.SelectedItem as TabItem;
                        clist = VisualTreeEnumHelper.GetAllControls((Visual)x.Content);                        

                        foreach (Visual v in clist)
                            if (v is WebBrowser)
                            {
                                WebBrowser wb = v as WebBrowser;

                                wb.LoadCompleted += new LoadCompletedEventHandler(wbloadcompl);
                                
                                wb.Navigate(fullpath);

                                break;
                            }
                    }
                    reader.Close();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }//end get zvit

        private void wbloadcompl(object sender, NavigationEventArgs e)
        {
            TabItem x = this.TabPanel.SelectedItem as TabItem;
            List<Visual> clist = VisualTreeEnumHelper.GetAllControls((Visual)x.Content);
            Dictionary<string, string> sdata = new Dictionary<string, string>();

            string fullpath2 = System.IO.Path.GetFullPath("s.txt");
            MyDataReader.readSaveFile(fullpath2, sdata);

            List<string> sdkeys = new List<string>(sdata.Keys);
            foreach(string s in sdkeys)
                sdata[s] = sdata[s].Replace(" $$a ", "\n");

            foreach (Visual v in clist)
                if (v is WebBrowser)
                {
                    WebBrowser wb = v as WebBrowser;

                    IHTMLDocument2 doc = wb.Document as IHTMLDocument2;
                    foreach (IHTMLElement el in doc.all)
                    {
                        if (el.getAttribute("name") != null)
                            if (sdata.ContainsKey(Convert.ToString(el.getAttribute("name"))))
                            {
                                el.setAttribute("value", sdata[el.getAttribute("name")]);
                            }
                    }
                    break;
                }
        }

        //close current tab
        private void CloseTab_Click(object sender, RoutedEventArgs e)
        {
            if (TabPanel.SelectedItem != null)
            {
                TabItem x = TabPanel.SelectedItem as TabItem;
                List<Visual> clist = VisualTreeEnumHelper.GetAllControls((Visual)x.Content);
                foreach (Visual v in clist)
                {
                    if (v is Label)
                    {
                        Label l = v as Label;
                        dataList.Remove(l.Content.ToString());
                        --counter;
                        break; // interrupt
                    }
                }
                TabPanel.Items.Remove(TabPanel.SelectedItem);
            }
        }

        //execute autofill from profile data
        private void ApplyProfileData(object sender, RoutedEventArgs e)
        {
            if (TabPanel.SelectedItem != null)
            {
                TabItem x = TabPanel.SelectedItem as TabItem;
                List<Visual> clist = VisualTreeEnumHelper.GetAllControls((Visual)x.Content);
                foreach (Visual v in clist)
                {
                    if (v is WebBrowser)
                    {
                        profile.applyProfileToWebbrowser(v as WebBrowser);
                        break; // interrupt
                    }
                }
            }
        }

        //GenerateXML
        private void GenerateXml(object sender, RoutedEventArgs e)
        {
            if (TabPanel.SelectedItem != null)
            {
                Dictionary<string, string> data = new Dictionary<string, string>();
                string header = "";
                TabItem x = TabPanel.SelectedItem as TabItem;
                List<Visual> clist = VisualTreeEnumHelper.GetAllControls((Visual)x.Content);
                foreach (Visual v in clist)
                {
                    if (v is Label)
                    {
                        Label l = v as Label;
                        header = l.Content.ToString();
                    }
                }

                string searchID = header.Substring(0, header.IndexOf(" ")); // ???

                foreach (Visual v in clist)
                {
                    if (v is WebBrowser)
                    {
                        MyDataReader.readWebpageData(v as WebBrowser, data, searchID);
                        //break; // interrupt
                    }
                }

                Fabric f = null; // = new z1_S0210110F();
                switch (searchID)
                {
                    case "z1": { f = new z1_S0210110F(); break; }
                }

                ZvitGeneral zg = f.create(profile, data);
                XmlSerializer serializer = null;

                string docdata = "";
                switch (f.zid)
                {
                    case "z1": { serializer = new XmlSerializer(typeof(z1_S0210110)); docdata = "S0210110"; break; }
                }


                string thename = profile.C_REG + profile.C_RAJ + "00" + profile.EDRPOU + docdata + "1" + "00" + data["CNT"]
                    + DateTime.Now.ToString("MMyyyy");
                using (var stream = new StreamWriter(thename + ".xml")) // path
                    serializer.Serialize(stream, zg);

            }  
        }

        //save file
        private void SaveFileClick(object sender, RoutedEventArgs e)
        {
            if (TabPanel.SelectedItem != null)
            {
                Dictionary<string, string> xdata = new Dictionary<string, string>();
                SaveFileWindow sf = new SaveFileWindow();

                TabItem x = this.TabPanel.SelectedItem as TabItem;
                List<Visual> clist = VisualTreeEnumHelper.GetAllControls((Visual)x.Content);
                foreach (Visual v in clist)
                {
                    if (v is WebBrowser)
                    {
                        WebBrowser wb = v as WebBrowser;
                        sf.hdoc = wb.Document as HTMLDocument;
                        MyDataReader.getWebPageData(v as WebBrowser, xdata);
                        sf.data = xdata;
                    }
                    if (v is Label)
                    {
                        Label l = v as Label;
                        sf.zvitid = l.Content.ToString();
                        sf.zvitid = sf.zvitid.Substring(0, sf.zvitid.IndexOf(" "));
                    }
                }
                sf.Show();
            }
        }

        //load file
        private void LoadFileClick(object sender, RoutedEventArgs e)
        {
            LoadFileWindow lf = new LoadFileWindow();
            lf.Show();
            lf.mf = this;
        }

        //check validation
        private void Validation_Click(object sender, RoutedEventArgs e)
        {
            if (TabPanel.SelectedItem != null)
            {
                TabItem x = this.TabPanel.SelectedItem as TabItem;
                List<Visual> clist = VisualTreeEnumHelper.GetAllControls((Visual)x.Content);

                WebBrowser wb = null;
                string id = "";
                foreach (Visual v in clist)
                {
                    if (v is WebBrowser)
                    {
                        wb = v as WebBrowser;
                    }
                    if (v is Label)
                    {
                        Label l = v as Label;
                        id = l.Content.ToString();
                        id = id.Substring(0, id.IndexOf(" "));
                    }
                }

                switch (id)
                {
                    case "z1": { MyValidation.z1(wb); break; }
                }
            }
        }

        //create word
        private void CreateWord(object sender, RoutedEventArgs e)
        {
            if (TabPanel.SelectedItem != null)
            {
                TabItem x = this.TabPanel.SelectedItem as TabItem;
                List<Visual> clist = VisualTreeEnumHelper.GetAllControls((Visual)x.Content);

                WebBrowser wb = null;
                string id = "";
                foreach (Visual v in clist)
                {
                    if (v is WebBrowser)
                    {
                        wb = v as WebBrowser;
                    }
                    if (v is Label)
                    {
                        Label l = v as Label;
                        id = l.Content.ToString();
                        id = id.Substring(0, id.IndexOf(" "));
                    }
                }
                MyWordEditor.GetDoc(wb, id);
            }
        }//end fun

        //UPDATERS

        //update HTML
        private void UpdateHtmlButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                byte[] bytes = File.ReadAllBytes("f:/style1.html");
                using (OleDbConnection connection = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = EsDb.mdb;"))
                {
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "UPDATE dovZvitForm SET zvitView = ? Where zvitShortName='z1'";
                    cmd.Parameters.AddWithValue("@zvitView", bytes);

                    cmd.Connection = connection;
                    try
                    {
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("An Item has been successfully updated", "Caption");
                        connection.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //update word
        private void UpdateWordButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                byte[] bytes = File.ReadAllBytes("f:/zvit001.doc");
                using (OleDbConnection connection = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = EsDb.mdb;"))
                {
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "UPDATE dovZvitForm SET zvitPrint = ? Where zvitShortName='z1'";
                    cmd.Parameters.AddWithValue("@zvitPrint", bytes);

                    cmd.Connection = connection;
                    try
                    {
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("An Item has been successfully updated", "Caption");
                        connection.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // 90% 10%
        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            SpritePanel.Height = DockElement.ActualHeight * 0.125;
            TabPanel.Height = DockElement.ActualHeight * 0.875;         
        }

        private void Form_Closed(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }
    }



}

