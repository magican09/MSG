using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class VOVRWork : Work
    {

        public const int VOVR_NUMBER_COL = MSGWork.MSG_NEEDS_OF_MACHINE_QUANTITY_COL + 1;
        public const int VOVR_NAME_COL = VOVR_NUMBER_COL + 1;
        public const int VOVR_MEASURE_COL = VOVR_NUMBER_COL + 2;
        public const int VOVR_QUANTITY_COL = VOVR_NUMBER_COL + 3;
        public const int VOVR_QUANTITY_FACT_COL = VOVR_NUMBER_COL + 4;
        public const int VOVR_LABOURNESS_COL = VOVR_NUMBER_COL + 5;


        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public override Excel.Worksheet Worksheet
        {
            get { return base.Worksheet; }
            set
            {
                this.KSWorks.Worksheet = value;
                this.WorkersComposition.Worksheet = value;
                this.MachinesComposition.Worksheet = value;
                base.Worksheet= value;

            }
        }

        private AdjustableCollection<KSWork> _kSWorks = new AdjustableCollection<KSWork>();

        [NonRegisterInUpCellAddresMap]
        public AdjustableCollection<KSWork> KSWorks
        {
            get { return _kSWorks; }
            set { SetProperty(ref _kSWorks, value); }
        }

        public VOVRWork() : base()
        {
            this.KSWorks.Owner = this;
        }

        public override void SetSectionNumber(string section_number)
        {
            base.SetSectionNumber(section_number);
            foreach (KSWork work in this.KSWorks)
            {
                work.SetSectionNumber(section_number);
            }

        }
        public override void UpdateExcelRepresetation()
        {
            VOVRWork vovr_work = this;
            this.UpdateExellBindableObject();
            foreach (KSWork ks_work in vovr_work.KSWorks.OrderBy(w => w.Number))
                ks_work.UpdateExcelRepresetation();

        }

        public override int AdjustExcelRepresentionTree(int row)
        {
            VOVRWork vovr_work = this;
            int vovr_row = row;
            vovr_work.ChangeTopRow(vovr_row);
            int ks_row = vovr_row;
            foreach (KSWork ks_work in vovr_work.KSWorks.OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberPrefix}.", ""))))
                ks_row = ks_work.AdjustExcelRepresentionTree(ks_row); ;

            vovr_row = ks_row;
            return vovr_row;
        }

        public override void SetStyleFormats(int col)
        {
            VOVRWork vovr_work = this;
            int vovr_work_col = col;
            var vovr_work_range = vovr_work.GetRange();
            vovr_work_range.Interior.ColorIndex = vovr_work_col;
            vovr_work_range.SetBordersLine();
            vovr_work.KSWorks.GetRange().SetBordersLine(XlLineStyle.xlLineStyleNone, XlLineStyle.xlDashDot, XlLineStyle.xlLineStyleNone, XlLineStyle.xlLineStyleNone);
            int ks_work_col = vovr_work_col;
            if (vovr_work.KSWorks.Count > 0)
            {
                Excel.Range _works_left_edge_range = vovr_work.KSWorks.Worksheet.Range[vovr_work.KSWorks[0]["Number"].Cell,
                                                                            vovr_work.KSWorks[vovr_work.KSWorks.Count - 1]["Number"].Cell];
                _works_left_edge_range.SetBordersLine(XlLineStyle.xlLineStyleNone, XlLineStyle.xlDashDot, XlLineStyle.xlLineStyleNone, XlLineStyle.xlLineStyleNone);

                foreach (KSWork ks_work in vovr_work.KSWorks)
                    ks_work.SetStyleFormats(ks_work_col);
            }


            //try
            //{
            //    Excel.Range top_row = this.RegisterSheet.Rows[vovr_work.KSWorks.GetTopRow() + 1];
            //    Excel.Range rottom_row_num = this.RegisterSheet.Rows[vovr_work.KSWorks.OrderBy(w => w.RCWorks.GetBottomRow()).Last().RCWorks.GetBottomRow()]; ;
            //    this.RegisterSheet.Range[top_row, rottom_row_num].Group();
            //}
            //catch { }
        }

        public override Range GetRange()
        {
            Excel.Range range = base.GetRange();
            Excel.Range ks_works_range = this.KSWorks.GetRange();
            range = Worksheet.Application.Union(new List<Excel.Range>() { range, ks_works_range });
            return range;
        }

        public override object Clone()
        {
            VOVRWork new_work = (VOVRWork)base.Clone();
            new_work.KSWorks = (AdjustableCollection<KSWork>)this.KSWorks.Clone();
            new_work.KSWorks.Owner = new_work;
            return new_work;
        }
    }
}
