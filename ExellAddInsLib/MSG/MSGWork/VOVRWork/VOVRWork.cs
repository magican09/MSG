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
                base.Worksheet = value;

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
            if (vovr_work.IsValid)
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
            //  Excel.Range ks_works_range = this.KSWorks.GetRange();
            //   range = Worksheet.Application.Union(new List<Excel.Range>() { range, ks_works_range });
            return range;
        }

        public override int GetLastRow()
        {
            int top_row = this.GetTopRow();
            int bottom_row = base.GetLastRow();
            int last_row = this.KSWorks.GetLastRow();
            if (last_row < bottom_row) last_row = bottom_row;

            return last_row;
        }
        public override object Clone()
        {
            VOVRWork new_work = (VOVRWork)base.Clone();
            new_work.KSWorks = (AdjustableCollection<KSWork>)this.KSWorks.Clone();
            new_work.KSWorks.Owner = new_work;
            return new_work;
        }
        public override void Validate()
        {
            decimal ks_laboriosness_sum = 0;
            foreach (var rc_work in this.KSWorks)
            {
                ks_laboriosness_sum += rc_work.Laboriousness * rc_work.ProjectQuantity;

            }
            var curent_work_laboriousness = this.Laboriousness * this.ProjectQuantity;
            if (Math.Round(ks_laboriosness_sum, 4) != Math.Round(curent_work_laboriousness, 4))
            {

            }

            bool is_valid = Math.Round(ks_laboriosness_sum, 3) == Math.Round(curent_work_laboriousness, 3);
            foreach (var ks_work in this.KSWorks)
            {
                ks_work.SetPropertyValidStatus("Laboriousness", is_valid);
                ks_work.SetPropertyValidStatus("ProjectQuantity", is_valid);
                ks_work.IsValid = is_valid;
            }

            this.KSWorks.Validate();
            base.Validate();
        }
        public void AddDeafaultChildWork(MSGExellModel model)
        {
            if (this.KSWorks.Count == 0)
            {
                var vovr_work = this;
                KSWork ks_work = new KSWork();
                ks_work.Worksheet = model.RegisterSheet;
                int rowIndex = vovr_work["Number"].Row;
                model.Register(ks_work, "Number", rowIndex, KSWork.KS_NUMBER_COL, model.RegisterSheet);
                model.Register(ks_work, "Code", rowIndex, KSWork.KS_CODE_COL, model.RegisterSheet);
                model.Register(ks_work, "Name", rowIndex, KSWork.KS_NAME_COL, model.RegisterSheet);
                model.Register(ks_work, "ProjectQuantity", rowIndex, KSWork.KS_QUANTITY_COL, model.RegisterSheet);
                model.Register(ks_work, "Quantity", rowIndex, KSWork.KS_QUANTITY_FACT_COL, model.RegisterSheet);
                model.Register(ks_work, "Laboriousness", rowIndex, KSWork.KS_LABOURNESS_COL, model.RegisterSheet);
                model.Register(ks_work, "UnitOfMeasurement.Name", rowIndex, KSWork.KS_MEASURE_COL, model.RegisterSheet);

                ks_work.Number = $"{vovr_work.Number}.1";
                ks_work.Code = "-";
                ks_work.Name = vovr_work.Name;
                ks_work.UnitOfMeasurement = vovr_work.UnitOfMeasurement;
                ks_work.ProjectQuantity = vovr_work.ProjectQuantity;
                ks_work.Laboriousness = vovr_work.Laboriousness;

                ks_work.AddDeafaultChildWork(model);
                this.KSWorks.Add(ks_work);
            }
        }

    }
}
