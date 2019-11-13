using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using mshtml;

namespace EzStatWPF
{
    public class CProfile
    {
        public string name;
        public string FIRM_NAME;
        public string FIRM_ADR;
        public string FIRM_ADR_FIZ;
        public string EDRPOU;
        public string VIK;
        public string VIK_RUK;
        public string VIK_TEL;
        public string VIK_EMAIL;
        public string FIRM_FAXORG;
        public string C_REG;
        public string C_RAJ;
        public string FIRM_SPATO;
        public string FIRM_KVED;
        public string FIZ_O;

        public void applyProfileToWebbrowser(WebBrowser wb)
        {
            IHTMLDocument2 doc = wb.Document as IHTMLDocument2;
            foreach (IHTMLElement el in doc.all)
            {
                switch (el.getAttribute("name"))
                {
                    //EDRPOU
                    case "E1": { el.setAttribute("value", EDRPOU.Substring(0, 1)); break; }
                    case "E2": { el.setAttribute("value", EDRPOU.Substring(1, 1)); break; }
                    case "E3": { el.setAttribute("value", EDRPOU.Substring(2, 1)); break; }
                    case "E4": { el.setAttribute("value", EDRPOU.Substring(3, 1)); break; }
                    case "E5": { el.setAttribute("value", EDRPOU.Substring(4, 1)); break; }
                    case "E6": { el.setAttribute("value", EDRPOU.Substring(5, 1)); break; }
                    case "E7": { el.setAttribute("value", EDRPOU.Substring(6, 1)); break; }
                    case "E8": { el.setAttribute("value", EDRPOU.Substring(7, 1)); break; }

                    case "FIRM_NAME": { el.setAttribute("value", FIRM_NAME);  break; }
                    case "FIRM_ADR": { el.setAttribute("value", FIRM_ADR); break; }
                    case "FIRM_ADR_FIZ": { el.setAttribute("value",FIRM_ADR_FIZ); break; }
                    case "MY_DATE": { el.setAttribute("value", DateTime.Today.ToString("dd.MM.yyyy")); break; }
                    case "VIK_RUK": { el.setAttribute("value", VIK_RUK); break; }
                    case "VIK": { el.setAttribute("value", VIK); break; }
                    case "VIK_TEL": { el.setAttribute("value", VIK_TEL); break; }
                    case "FIRM_FAXOR": { el.setAttribute("value", FIRM_FAXORG); break; }
                    case "FIRM_FAXORG": { el.setAttribute("value", FIRM_FAXORG); break; }
                    case "VIK_EMAIL": { el.setAttribute("value", VIK_EMAIL); break; }
                }
               
            }

        }

    }
}
