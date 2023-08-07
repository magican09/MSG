using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public static class Func
    {
        public static bool isValidGUID(string str)
        {
            string strRegex = @"^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$";
            Regex re = new Regex(strRegex);
            if (re.IsMatch(str))
                return (true);
            else
                return (false);
        }
        public static bool containValidGUID(string str)
        {
            string strRegex = @"^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$";
            Regex re = new Regex(strRegex);

            if (re.Matches(str).Count>0)
                return (true);
            else
                return (false);
        }
        public static List<string> getGUIDs(string str)
        {
            string strRegex = @"^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$";
            Regex re = new Regex(strRegex);
            List<string> out_list = new List<string>();
            foreach(Match match in re.Matches(str))
                out_list.Add(match.Value);
            return out_list;
               
        }
        public static string RangeAddress(Excel.Range rng)
        {
            return rng.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1,
                   Type.Missing, Type.Missing);
        }

        private static bool IsSubscribed(EventHandler @event, object evHandler)
        {
            return @event.GetInvocationList().Contains(evHandler);
        }
    }

 }


