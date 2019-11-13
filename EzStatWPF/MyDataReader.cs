using mshtml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace EzStatWPF
{
    public class MyDataReader
    {
        public static void readWebpageData(WebBrowser wb, Dictionary<string, string> d, string id)
        {
            int period = 0;
            using (OleDbConnection connection = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = EsDb.mdb;"))
            {
                OleDbCommand commandGetPeriod = new OleDbCommand("Select zvitPeriod From dovZvitForm Where zvitShortName='" + id + "'");
                commandGetPeriod.Connection = connection;
                try
                {
                    connection.Open();
                    OleDbDataReader reader = commandGetPeriod.ExecuteReader();
                    while (reader.Read())
                    {
                        period = Convert.ToInt32(reader[0]);
                    }
                    reader.Close();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

          
            d.Add("CNT", returnCNT(id,period)); // some special data from DB

            getWebPageData(wb, d);
        }//end fun

        public static void readSaveFile(string fullpath, Dictionary<string, string> d)
        {
            string[] lines = System.IO.File.ReadAllLines(fullpath);
            // Display the file contents by using a foreach loop.            
            foreach (string l in lines)
            {
                string id = l.Substring(3, l.IndexOf("val")-4);
                string val = l.Substring(l.IndexOf(" val ") + 5, l.Length - (l.IndexOf(" val ") + 5));
                d.Add(id, val);
            }
           
        }

        public static void getWebPageData(WebBrowser wb, Dictionary<string, string> d)
        {
            IHTMLDocument2 doc = wb.Document as IHTMLDocument2;
            //d.Clear(); // ??
            File.WriteAllText("test.html", doc.body.innerHTML);
            foreach (IHTMLElement el in doc.all)
            {
                if (el.getAttribute("type") != null)
                {
                    if (el.getAttribute("type").Equals("text"))
                        if (el.getAttribute("value") != null)
                        {
                            
                            d.Add(el.getAttribute("name"), el.getAttribute("value")); }
                        else {
                            
                            d.Add(el.getAttribute("name"), ""); }
                }

                if (el.outerHTML != null)
                    if (el.outerHTML.Length > 10)
                        if (el.outerHTML.Substring(0, 9).Equals("<textarea"))
                            d.Add(el.getAttribute("name"), el.getAttribute("value"));

            }
        }

        private static string returnCNT(string zid, int zperiod)
        {
            int cnt = 1;
            List<string> datelist = new List<string>();
            List<int> cntlist = new List<int>();

            using (OleDbConnection connection = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = EsDb.mdb;"))
            {                
                OleDbCommand commandGetCNT = new OleDbCommand("Select * From ZvitReg Where docId='" + zid + "'"+
                    " and zvitPeriod="+zperiod);
                commandGetCNT.Connection = connection;
                try
                {
                    connection.Open();
                    OleDbDataReader reader = commandGetCNT.ExecuteReader();
                    while (reader.Read())
                    {
                        datelist.Add(reader[1].ToString());
                        cntlist.Add(Convert.ToInt32(reader[3]));
                    }
                    reader.Close();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            if(datelist.Count !=0)
            switch(zperiod)
            {
                case 1: {
                        string currmonth = DateTime.Now.Month.ToString();
                        
                        for(int i=0;i<datelist.Count;++i)
                        {
                            if (datelist[i].Substring(3, 2).Equals(currmonth))
                            { if (cnt < cntlist[i]) cnt = cntlist[i]; }
                        }
                            ++cnt;
                            break; }//1 month
                case 2: {
                            string currmonth = DateTime.Now.Month.ToString();
                            int quartal = 0;
                            if (currmonth.Equals("01") || currmonth.Equals("02") || currmonth.Equals("03"))
                                quartal = 1;
                            else
                            if (currmonth.Equals("04") || currmonth.Equals("05") || currmonth.Equals("06"))
                                quartal = 2;
                            else
                            if (currmonth.Equals("07") || currmonth.Equals("08") || currmonth.Equals("09"))
                                quartal = 3;
                            else
                            if (currmonth.Equals("10") || currmonth.Equals("11") || currmonth.Equals("12"))
                                quartal = 4;

                            for (int i = 0; i < datelist.Count; ++i)
                            {
                                if (quartal == 1)
                                    if (datelist[i].Substring(3, 2).Equals("01") || datelist[i].Substring(3, 2).Equals("02") || datelist[i].Substring(3, 2).Equals("03"))
                                    { if (cnt < cntlist[i]) cnt = cntlist[i]; }
                                    else if(quartal == 2)
                                        if (datelist[i].Substring(3, 2).Equals("04") || datelist[i].Substring(3, 2).Equals("05") || datelist[i].Substring(3, 2).Equals("06"))
                                        { if (cnt < cntlist[i]) cnt = cntlist[i]; }
                                else if(quartal == 3)
                                            if (datelist[i].Substring(3, 2).Equals("07") || datelist[i].Substring(3, 2).Equals("08") || datelist[i].Substring(3, 2).Equals("09"))
                                            { if (cnt < cntlist[i]) cnt = cntlist[i]; }
                                else if(quartal == 4)
                                                if (datelist[i].Substring(3, 2).Equals("10") || datelist[i].Substring(3, 2).Equals("11") || datelist[i].Substring(3, 2).Equals("12"))
                                                { if (cnt < cntlist[i]) cnt = cntlist[i]; }
                            }
                            ++cnt;
                            break; }//3month
                case 3: {
                            string currmonth = DateTime.Now.Month.ToString();
                            int quartal = 0;
                            if (currmonth.Equals("01") || currmonth.Equals("02") || currmonth.Equals("03")
                                || currmonth.Equals("04") || currmonth.Equals("05") || currmonth.Equals("06"))
                                quartal = 1;
                            else if (currmonth.Equals("07") || currmonth.Equals("08") || currmonth.Equals("09")
                                || currmonth.Equals("10") || currmonth.Equals("11") || currmonth.Equals("12"))
                                quartal = 2;

                            for (int i = 0; i < datelist.Count; ++i)
                            {
                                if (quartal == 1)
                                    if (datelist[i].Substring(3, 2).Equals("01") || datelist[i].Substring(3, 2).Equals("02") || datelist[i].Substring(3, 2).Equals("03")
                                        || datelist[i].Substring(3, 2).Equals("04") || datelist[i].Substring(3, 2).Equals("05") || datelist[i].Substring(3, 2).Equals("06"))
                                    { if (cnt < cntlist[i]) cnt = cntlist[i]; }
                                    else if (quartal == 2)
                                        if (datelist[i].Substring(3, 2).Equals("07") || datelist[i].Substring(3, 2).Equals("08") || datelist[i].Substring(3, 2).Equals("09")
                                            || datelist[i].Substring(3, 2).Equals("10") || datelist[i].Substring(3, 2).Equals("11") || datelist[i].Substring(3, 2).Equals("12"))
                                        { if (cnt < cntlist[i]) cnt = cntlist[i]; }                                       
                            }
                            ++cnt;
                            break; }//6 month
                case 4: {
                            string currmonth = DateTime.Now.Month.ToString();
                            for (int i = 0; i < datelist.Count; ++i)
                            {                              
                               if (datelist[i].Substring(3, 2).Equals("01") || datelist[i].Substring(3, 2).Equals("02") || datelist[i].Substring(3, 2).Equals("03")
                                    || datelist[i].Substring(3, 2).Equals("04") || datelist[i].Substring(3, 2).Equals("05") || datelist[i].Substring(3, 2).Equals("06")
                                    || datelist[i].Substring(3, 2).Equals("07") || datelist[i].Substring(3, 2).Equals("08") || datelist[i].Substring(3, 2).Equals("09"))
                                    { if (cnt < cntlist[i]) cnt = cntlist[i]; }
                                   
                            }
                            ++cnt;
                            break; }//9 month
                case 5: {
                            string curryear = DateTime.Now.Year.ToString();

                            for (int i = 0; i < datelist.Count; ++i)
                            {
                                if (datelist[i].Substring(6, 4).Equals(curryear))
                                { if (cnt < cntlist[i]) cnt = cntlist[i]; }
                            }
                            ++cnt;
                            break; }//1 year
            }
            //push new CNT !!!!!!
            pushCNT(zid, zperiod, cnt);

            if (cnt >= 10000)
                return cnt.ToString();
            else if (cnt >= 1000)
                return "0" + cnt;
            else if (cnt >= 100)
                return "00" + cnt;
            else if (cnt >= 10)
                return "000" + cnt;
            else return "0000" + cnt;
        }

        private static void pushCNT(string zid, int zperiod, int newcnt)
        {
            string cs = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=EsDb.mdb;";
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand cmd = new OleDbCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "insert into ZvitReg ([zvitPeriod],[data],[docId],[c_doc_cnt]) values (?,?,?,?)";
            cmd.Parameters.AddWithValue("@zvitPeriod", zperiod);
            cmd.Parameters.AddWithValue("@data", DateTime.Now.ToString("dd.MM.yyyy"));
            cmd.Parameters.AddWithValue("@docId", zid);
            cmd.Parameters.AddWithValue("@c_doc_cnt", newcnt);
           
            cmd.Connection = con;
            try
            {
                con.Open();
                cmd.ExecuteNonQuery();           
                con.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
