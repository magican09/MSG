using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG.Section
{
    public class WorksSection : ExcelBindableBase
    {
        public const int WSEC_NUMBER_COL = 2;
        public const int WSEC_NAME_COL = WSEC_NUMBER_COL + 1;
        public const int _SECTIONS_GAP = 2;
        public const int _MSG_WORKS_GAP = 1;
        private Excel.Worksheet _worksheet;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public override Excel.Worksheet Worksheet
        {
            get { return _worksheet; }
            set
            {
                _worksheet = value;
                this.MSGWorks.Worksheet = _worksheet;
            }
        }


        private string _number;

        public override string Number
        {
            get { return _number; }
            set { SetProperty(ref _number, value); }
        }//Номер работы

        private string _name;

        public string Name
        {
            get { return _name; }
            set { SetProperty(ref _name, value); }
        }//Наименование работы
        /// <summary>
        /// Коллекция с работами типа МСГ модели
        /// </summary>
        private AdjustableCollection<MSGWork> _mSGWorks = new AdjustableCollection<MSGWork>();

        [NonRegisterInUpCellAddresMap]
        public AdjustableCollection<MSGWork> MSGWorks
        {
            get { return _mSGWorks; }
            set { _mSGWorks = value; }
        }
        public override void UpdateExcelRepresetation()
        {
            WorksSection w_section = this;
            this.UpdateExellBindableObject();
            foreach (MSGWork msg_work in w_section.MSGWorks.OrderBy(w => w.Number))
                msg_work.UpdateExcelRepresetation();

        }
        public override int AdjustExcelRepresentionTree(int row)
        {
            int section_row = row;
            int msg_row = row - _MSG_WORKS_GAP;
            var w_section = this;
            w_section.ChangeTopRow(section_row);
            foreach (MSGWork msg_work in w_section.MSGWorks.OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberPrefix}.", ""))))
                msg_row = msg_work.AdjustExcelRepresentionTree(msg_row + _MSG_WORKS_GAP);

            section_row = msg_row;
            return section_row;
        }


        public override void SetStyleFormats(int col)
        {
            WorksSection section = this;
            int selectin_col = col;
            var section_range = section.GetRange(WSEC_NAME_COL);
            section_range.Interior.ColorIndex = selectin_col;
            section_range.SetBordersLine();
            int first_row = section.GetTopRow();
            if (section.MSGWorks.Count > 0)
            {
                Excel.Range msg_works_left_edge_range = section.MSGWorks.Worksheet.Range[section.MSGWorks[0].CellAddressesMap["Number"].Cell, section.MSGWorks[section.MSGWorks.Count - 1].CellAddressesMap["Number"].Cell];
                msg_works_left_edge_range.SetBordersLine(XlLineStyle.xlLineStyleNone, XlLineStyle.xlDashDot, XlLineStyle.xlLineStyleNone, XlLineStyle.xlLineStyleNone);

                int msg_work_col = selectin_col + 1;
                int last_section_row = 0;
                foreach (MSGWork msg_work in section.MSGWorks)
                    msg_work.SetStyleFormats(msg_work_col);

            }

            try
            {
                var section_full_range = section.GetRange();
                var lowest_edge_range = section_full_range.GetRangeWithLowestEdge();
                // Excel.Range range = Worksheet.Range[Worksheet.Rows[section_full_range.Row + 1], lowest_edge_range.Rows[lowest_edge_range.Rows.Count + _SECTIONS_GAP]];
                Excel.Range range = Worksheet.Range[Worksheet.Rows[section_full_range.Row + 1], lowest_edge_range.Rows[lowest_edge_range.Rows.Count + _SECTIONS_GAP]];
                range.Group();
            }
            catch
            {

            }

        }
        public WorksSection()
        {

            this.MSGWorks.Owner = this;
            this.MSGWorks.Worksheet = this.Worksheet;
        }

        public override Range GetRange()
        {
            Excel.Range range = base.GetRange();
            Excel.Range msg_works_range = this.MSGWorks.GetRange();
            range = Worksheet.Application.Union(new List<Excel.Range>() { range, msg_works_range });

            return range;
        }
        public override object Clone()
        {
            WorksSection new_obj = (WorksSection)base.Clone();
            new_obj.MSGWorks = (AdjustableCollection<MSGWork>)this.MSGWorks.Clone();
            new_obj.MSGWorks.Owner = new_obj;
            return new_obj;
        }

    }
}
