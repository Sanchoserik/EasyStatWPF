using mshtml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EzStatWPF
{
    public class MyValidation
    {
        private static string stylegreen = "green"; //background-color:
        private static string stylered = "red";

        public static void z1(System.Windows.Controls.WebBrowser wb)
        {
            HTMLDocument doc = wb.Document as HTMLDocument;
            Dictionary<string, decimal> zdata = new Dictionary<string, decimal>();
            //parse decimal
            string a1 = getElement(doc, "A010_1");
            a1 = a1.Replace(".", ",");
            string a2 = getElement(doc, "A010_2");
            a2 = a2.Replace(".", ",");
            string a3 = getElement(doc, "A010_3");
            a3 = a3.Replace(".", ",");
            string a4 = getElement(doc, "A010_4");
            a4 = a4.Replace(".", ",");
            string a5 = getElement(doc, "A010_5");
            a5 = a5.Replace(".", ",");
            string a6 = getElement(doc, "A010_6");
            a6 = a6.Replace(".", ",");

            try
            {
                decimal a010_1 = Decimal.Parse(a1);
                decimal a010_2 = Decimal.Parse(a2);
                decimal a010_3 = Decimal.Parse(a3);
                decimal a010_4 = Decimal.Parse(a4);
                decimal a010_5 = Decimal.Parse(a5);
                decimal a010_6 = Decimal.Parse(a6);

                zdata.Add("A010_1", a010_1);
                zdata.Add("A010_2", a010_2);
                zdata.Add("A010_3", a010_3);
                zdata.Add("A010_4", a010_4);
                zdata.Add("A010_5", a010_5);
                zdata.Add("A010_6", a010_6);

                bool validation = readscript(wb, "f:/script.txt", zdata);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Одне з числових полів має неправильний формат заповнення");
            }
        }//end fun


        static bool readscript(System.Windows.Controls.WebBrowser wb, string scr, Dictionary<string, decimal> data)
        {
            HTMLDocument doc = wb.Document as HTMLDocument;

            bool isvalid = true;
            bool lockvalid = false;

            using (StreamReader sr = new StreamReader(scr))
            {
                string line;
                // Read and display lines from the file until the end of 
                // the file is reached.
                while ((line = sr.ReadLine()) != null)
                {
                    string leftrarg = "", operation = "", rightarg = "";
                    string l = line;
                    decimal arg1 = 0, arg2 = 0;

                    for (int i = 0; i < l.Length; ++i)
                    {
                        if (l.Substring(i, 1).Equals(" "))
                        {
                            leftrarg = l.Substring(0, i);
                            if (l.Substring(i + 3, 1).Equals(" ")) { operation = l.Substring(i + 1, 2); i += 4; rightarg = l.Substring(i, l.Length - i); break; }
                            else { operation = l.Substring(i + 1, 1); i += 3; rightarg = l.Substring(i, l.Length - i); break; }
                        }
                    }

                    arg1 = data[leftrarg];
                    if (rightarg.Substring(0, 1).Equals("E"))
                    {
                        string s = "";
                        for (int i = 2; i < rightarg.Length; ++i)
                        {
                            if (!rightarg.Substring(i, 1).Equals(" "))
                                s += rightarg.Substring(i, 1);
                            else { arg2 += data[s]; s = ""; }
                        }
                    }
                    else
                        arg2 = data[rightarg];

                    //
                    switch (operation)
                    {
                        case ">":
                            {
                                if (arg1 > arg2) { if (!lockvalid) { isvalid = true; setStlyle(doc, leftrarg, stylegreen); } }
                                else { isvalid = false; lockvalid = true; setStlyle(doc, leftrarg, stylered); }
                                break;
                            }
                        case "<":
                            {
                                if (arg1 < arg2) { if (!lockvalid) { isvalid = true; setStlyle(doc, leftrarg, stylegreen); } }
                                else { isvalid = false; lockvalid = true; setStlyle(doc, leftrarg, stylered); }
                                break;
                            }
                        case "<=":
                            {
                                if (arg1 <= arg2) { if (!lockvalid) { isvalid = true; setStlyle(doc, leftrarg, stylegreen); } }
                                else { isvalid = false; lockvalid = true; setStlyle(doc, leftrarg, stylered); }
                                break;
                            }
                        case ">=":
                            {
                                if (arg1 >= arg2) { if (!lockvalid) { isvalid = true; setStlyle(doc, leftrarg, stylegreen); } }
                                else { isvalid = false; lockvalid = true; setStlyle(doc, leftrarg, stylered); }
                                break;
                            }
                        case "==":
                            {
                                if (arg1 == arg2) { if (!lockvalid) { isvalid = true; setStlyle(doc, leftrarg, stylegreen); } }
                                else { isvalid = false; lockvalid = true; setStlyle(doc, leftrarg, stylered); }
                                break;
                            }
                        case "!=":
                            {
                                if (arg1 != arg2) { if (!lockvalid) { isvalid = true; setStlyle(doc, leftrarg, stylegreen); } }
                                else { isvalid = false; lockvalid = true; setStlyle(doc, leftrarg, stylered); }
                                break;
                            }
                    }

                }//end while
            }
            return isvalid;
        }

        static void setStlyle(HTMLDocument doc , string id, string style)
        {
            // doc.documentElement.GetElementById(id).Style = style;
            IHTMLElement el = doc.getElementById(id); //.style = style;
           // el.setAttribute("style", style);
            el.style.backgroundColor = style;
        }

        static string getElement(HTMLDocument doc, string id)
        {
            IHTMLElement el = doc.getElementById(id);
            string val = el.getAttribute("value"); // 
            if (val == null) val = "";
            return val;
        }//end fun

    }
}
