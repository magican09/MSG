using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class ExellPropAddress
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public Excel.Worksheet Worksheet { get; set; }
        public string ProprertyName { get; set; }
        public Excel.Range Cell
        {
            get { return this.Worksheet.Cells[Row, Column]; }
        }
        public ExellPropAddress()
        {

        }
        public ExellPropAddress(int row, int column, Excel.Worksheet worksheet, string prop_name = "")//,)
        {
            Row = row;
            Column = column;
            Worksheet = worksheet;
            ProprertyName = prop_name;
            this.Cell.Interior.Color = XlRgbColor.rgbGreenYellow;
        }
    }
}
