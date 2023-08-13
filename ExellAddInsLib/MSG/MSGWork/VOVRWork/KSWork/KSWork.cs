using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class KSWork : Work
    {
        public const int KS_NUMBER_COL = 23;
        public const int KS_CODE_COL = KS_NUMBER_COL + 1;
        public const int KS_NAME_COL = KS_NUMBER_COL + 2;
        public const int KS_MEASURE_COL = KS_NUMBER_COL + 3;
        public const int KS_QUANTITY_COL = KS_NUMBER_COL + 4;
        public const int KS_QUANTITY_FACT_COL = KS_NUMBER_COL + 5;
        public const int KS_LABOURNESS_COL = KS_NUMBER_COL + 6;

        private Excel.Worksheet _worksheet;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public override Excel.Worksheet Worksheet
        {
            get { return _worksheet; }
            set
            {
                _worksheet = value;
                this.RCWorks.Worksheet = _worksheet;
                this.WorkersComposition.Worksheet = _worksheet;
                this.MachinesComposition.Worksheet = _worksheet;
                this.CellAddressesMap.SetWorksheet(_worksheet);
            }
        }

        private string _code;

        public string Code
        {
            get { return _code; }
            set { SetProperty(ref _code, value); }
        }
        private AdjustableCollection<RCWork> _rCWorks = new AdjustableCollection<RCWork>();

        [NonRegisterInUpCellAddresMap]
        public AdjustableCollection<RCWork> RCWorks
        {
            get { return _rCWorks; }
            set { SetProperty(ref _rCWorks, value); }
        }
        public KSWork() : base()
        {
            this.RCWorks.Owner = this;
        }
        public override void SetSectionNumber(string section_number)
        {
            base.SetSectionNumber(section_number);
            foreach (RCWork work in this.RCWorks)
            {
                work.SetSectionNumber(section_number);
                if (work.ReportCard != null)
                    work.ReportCard.Number = work.Number;
            }

        }
        public override void UpdateExcelRepresetation()
        {
            KSWork ks_work = this;
            this.UpdateExellBindableObject();
            foreach (RCWork rc_work in ks_work.RCWorks.OrderBy(w => w.Number))
            {
                rc_work.UpdateExcelRepresetation();
            }
        }


        public override int AdjustExcelRepresentionTree(int row)
        {
            KSWork ks_work = this;
            int ks_row = row;
            ks_work.ChangeTopRow(ks_row);
            //var duple_kc_works = this.KSWorks.Where(ksw => ksw.Number == ks_work.Number && ksw.Id != ks_work.Id).ToList();
            int ks_work_cuont = 0;
            //foreach (var ksw in duple_kc_works)
            //{
            //    ks_work_cuont++;
            //    ksw.ChangeTopRow(ks_row + ks_work_cuont);
            //}
            
            int rc_row = ks_row + ks_work_cuont;
            foreach (RCWork rc_work in ks_work.RCWorks.OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberPrefix}.", ""))))
            {
                rc_row = rc_work.AdjustExcelRepresentionTree( rc_row);
                rc_row++;
            }

            ks_row = rc_row;
            return ks_row;
        }

        public override void  SetStyleFormats(int row)
        {
            KSWork ks_work = this;
            int ks_work_col = row;
            var ks_work_range = ks_work.GetRange(KS_LABOURNESS_COL);
            ks_work_range.Interior.ColorIndex = ks_work_col;
            int last_row = ks_work_range.Row;
            Excel.Range rc_works_range =  ks_work.RCWorks.GetRange();
            if (rc_works_range == null)
                ks_work.RCWorks.GetRange();
            rc_works_range.SetBordersBoldLine(XlLineStyle.xlDouble);
            rc_works_range.Interior.ColorIndex = ks_work_col;
            ks_work_range.SetBordersBoldLine(XlLineStyle.xlLineStyleNone);
            foreach (RCWork rc_work in ks_work.RCWorks)
                if(rc_work.ReportCard!=null) rc_work.ReportCard.SetStyleFormats(ks_work_col);
           // ks_work_range.SetBordersBoldLine(XlLineStyle.xlLineStyleNone);
        }
        public override Range GetRange()
        {
            Excel.Range range = base.GetRange();
            Excel.Range rc_works_range = this.RCWorks.GetRange();
            range = Worksheet.Application.Union(new List<Excel.Range>() { range, rc_works_range });
            return range;
          
        }

        public override object Clone()
        {
            KSWork new_work = (KSWork)base.Clone();
            new_work.Code = Code;
            new_work.RCWorks = (AdjustableCollection<RCWork>)this.RCWorks.Clone();
            new_work.RCWorks.Owner = new_work;
            return new_work;
        }
    }
}
