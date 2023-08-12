using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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

            if (re.Matches(str).Count > 0)
                return (true);
            else
                return (false);
        }
        public static List<string> getGUIDs(string str)
        {
            string strRegex = @"^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$";
            Regex re = new Regex(strRegex);
            List<string> out_list = new List<string>();
            foreach (Match match in re.Matches(str))
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
     
        public static void SetBordersBoldLine(this Excel.Range range)
        {
            if (range == null) return;
            //range.Borders.LineStyle = Excel.XlLineStyle.xlDot;
            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlDouble;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlDouble;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDouble;
        }
        public static Excel.Range GetRangeWithLowestEdge(this Excel.Range range)
        {
            Excel.Range lowest_range = null;
            foreach (Excel.Range r in range)
            {
                if (lowest_range == null)
                    lowest_range = r;
                else if (lowest_range.Rows[lowest_range.Rows.Count].Row < r.Row)
                    lowest_range = r;
            }
            return lowest_range;
        }
        public static Excel.Range Union(this Excel._Application aplication,List<Excel.Range> ranges)
        {
            Excel.Range union_range = null;
            foreach(Excel.Range r in ranges.Where(r=>r!=null))
            {
                if (union_range == null) union_range = r;
                else
                    union_range = aplication.Union(union_range,r);
            }
            return union_range;
        }
            /// <summary>
            /// Функция устанавливает границы диапазона двойной линей
            /// </summary>
            /// <param name="range"></param>
            /// <param name="right"></param>
            /// <param name="left"></param>
            /// <param name="top"></param>
            /// <param name="bottom"></param>
            public static void SetBordersBoldLine(this Excel.Range range, bool right = true, bool left = true, bool top = true, bool bottom = true)
        {
            if (range == null) return;

            if (left) range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            if (top) range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            if (bottom) range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            if (right) range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }

        /// <summary>
        /// Функция устанавливает границы диапазона соовествующими типами линий
        /// </summary>
        /// <param name="range"></param>
        /// <param name="right"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="bottom"></param>
        public static  void SetBordersBoldLine(this Excel.Range range,
            Excel.XlLineStyle right = Excel.XlLineStyle.xlDouble,
            Excel.XlLineStyle left = Excel.XlLineStyle.xlDouble,
            Excel.XlLineStyle top = Excel.XlLineStyle.xlDouble,
            Excel.XlLineStyle bottom = Excel.XlLineStyle.xlDouble)
        {
            if (range == null) return;

            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = left;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = top;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = bottom;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = right;
        }
    }

}


