using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class VOVRWork : Work
    {
        private AdjustableCollection<KSWork> _kSWorks = new AdjustableCollection<KSWork>();
      
        private Excel.Worksheet _worksheet;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public override Excel.Worksheet Worksheet
        {
            get { return _worksheet; }
            set
            {
                _worksheet = value;
                this.KSWorks.Worksheet = _worksheet;
                this.WorkersComposition.Worksheet = _worksheet;
                this.MachinesComposition.Worksheet = _worksheet;
                this.CellAddressesMap.SetWorksheet(_worksheet);

            }
        }


        [NonRegisterInUpCellAddresMap]
        public AdjustableCollection<KSWork> KSWorks
        {
            get { return _kSWorks; }
            set { SetProperty(ref _kSWorks, value); }
        }

        public VOVRWork() : base()
        {

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
      
        public override  int AdjustExcelRepresentionTree( int row)
        {
            VOVRWork vovr_work = this;
            int vovr_row = row;
            vovr_work.ChangeTopRow(vovr_row);
            //var duple_vovr_works = this.VOVRWorks.Where(vrw => vrw.Number == vovr_work.Number && vrw.Id != vovr_work.Id).ToList();
            //int vovr_work_cuont = 0;
            //foreach (var vrw in duple_vovr_works)
            //{
            //    vovr_work_cuont++;
            //    vrw.ChangeTopRow(vovr_row + vovr_work_cuont);
            //}
            int ks_row = vovr_row;
            foreach (KSWork ks_work in vovr_work.KSWorks.OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberSuffix}.", ""))))
            {
                ks_row = ks_work.AdjustExcelRepresentionTree( ks_row); ;
            }

            vovr_row = ks_row;
            return vovr_row;
        }

        public override void  SetStyleFormats( int col)
        {
            VOVRWork vovr_work = this;
            int vovr_work_col = col;
            var vovr_work_range = vovr_work.GetRange();
            vovr_work_range.Interior.ColorIndex = vovr_work_col;
            vovr_work_range.SetBordersBoldLine();
            vovr_work.KSWorks.GetRange().SetBordersBoldLine( XlLineStyle.xlLineStyleNone, XlLineStyle.xlDashDot, XlLineStyle.xlLineStyleNone, XlLineStyle.xlLineStyleNone);
            int ks_work_col = vovr_work_col;
            int last_vovr_row = vovr_work_range.Row;
            foreach (KSWork ks_work in vovr_work.KSWorks)
                ks_work.SetStyleFormats(ks_work_col);
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
