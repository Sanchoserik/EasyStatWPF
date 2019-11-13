using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using mshtml;

namespace EzStatWPF
{
    class MyWordEditor
    {

        public static void GetDoc(WebBrowser wb, string zid)
        {
            using (OleDbConnection connection = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = EsDb.mdb;"))
            {
                OleDbCommand command = new OleDbCommand("Select zvitPrint From dovZvitForm Where zvitShortName='" + zid + "'");
                command.Connection = connection;
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        string outpath = "docFile" + zid + ".doc";
                        File.WriteAllBytes(outpath, (byte[])reader[0]);

                        switch (zid)
                        {
                            case "z1": { z1doc(wb, outpath); break; }

                        }

                    }
                    reader.Close();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }//end fun
        }


        //find and replace in word
        private static void FindAndReplace(Microsoft.Office.Interop.Word.Application fileOpen, object findText, object replaceWithText)
        {
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            if (replaceWithText.ToString().Length > 10)
            {

                string oldline = replaceWithText.ToString();
                replaceWithText = "-R-";

                fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                       ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                       ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);


                replaceWithText = oldline;
                findText = "-R-";

                while (oldline.Length > 10)
                {
                    string subline = oldline.Substring(0, 10) + "-R-";
                    oldline = oldline.Remove(0, 10);

                    replaceWithText = subline;
                    //execute find and replace
                    fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                        ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                        ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);

                }//end while
                replaceWithText = oldline;
                fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                      ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                      ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            }
            else fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                       ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                       ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        private static void wordTableOptimizer(HTMLDocument doc, int colCount, int wordIndex, int htmltableindex, string path, Microsoft.Office.Interop.Word.Document wDoc)
        {
            if (doc.getElementById("edittable" + htmltableindex) != null)  //&& doc.getElementById("edittable").innerHTML.IndexOf("RowX") != -1)
            {
                IHTMLElement el = doc.getElementById("edittable" + htmltableindex);
                string s = el.innerHTML;
                int counter = 0;
                while (true)
                {
                    if (s.IndexOf("XRow") != -1)
                    { ++counter; s = s.Substring(s.IndexOf("XRow") + 4); }
                    else break;
                }


                for (int i = 1; i <= counter; ++i)
                {
                    wDoc.Tables[wordIndex].Rows.Add();
                    for (int j = 1; j <= colCount; ++j) //???
                    {
                        wDoc.Tables[wordIndex].Cell(wDoc.Tables[wordIndex].Rows.Count, j).Range.Text = doc?.getElementById("TR" + j + "X" + i)?.getAttribute("value");
                    }
                }

            }
        }

        private static void z1doc(System.Windows.Controls.WebBrowser wb, string path)
        {
            Microsoft.Office.Interop.Word.Application fileOpen = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wDoc = fileOpen.Documents.Open(Path.GetFullPath(path), ReadOnly: false);
            wDoc.Activate();

            HTMLDocument doc = wb.Document as HTMLDocument;
            wordTableOptimizer(doc, 8, 6, 0, path, wDoc);

            Dictionary<string, string> wData = new Dictionary<string, string>();
            foreach (IHTMLElement el in doc.all)
            {
                if (el.getAttribute("type") != null && el.getAttribute("name") != null && el.getAttribute("value") != null)
                    if (el.getAttribute("type").Equals("text"))
                        wData.Add(el.getAttribute("name"), el.getAttribute("value"));
            }
            List<string> keys = new List<string>(wData.Keys);
            foreach (string s in keys)
                wData[s] = wData[s].Replace("\n", " ");

            //get text area value! and replace '\n'
            string textareaData = doc.getElementById("T2RXXXXG1S").getAttribute("value");
            textareaData = textareaData.Replace("\n", " ");
            FindAndReplace(fileOpen, "T2RXXXXG1S", textareaData);

            //replace by special keys
            replaceBySpecialKey(fileOpen, wData);

            foreach (KeyValuePair<string, string> kp in wData)
            {
                if( !kp.Key.Equals("FIRM_ADR_FIZ") && !kp.Key.Equals("VIK_RUK") && !kp.Key.Equals("VIK_TEL") && !kp.Key.Equals("VIK_EMAIL")) 
                FindAndReplace(fileOpen, kp.Key, kp.Value);
            }

            wDoc.Save();
            fileOpen.Quit();
        }

        private static void replaceBySpecialKey(Microsoft.Office.Interop.Word.Application fileOpen, Dictionary<string,string> data)
        {
            FindAndReplace(fileOpen, "FIRM_ADR_FIZ", data["FIRM_ADR_FIZ"]);
            FindAndReplace(fileOpen, "VIK_RUK", data["VIK_RUK"]);
            FindAndReplace(fileOpen, "VIK_TEL", data["VIK_TEL"]);
            FindAndReplace(fileOpen, "VIK_EMAIL", data["VIK_EMAIL"]);
        }

    }//end class
}
