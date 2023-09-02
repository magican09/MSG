using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class RCWork : Work
    {

        public const int RC_NUMBER_COL = KSWork.KS_LABOURNESS_COL + 1;
        public const int RC_CODE_COL = RC_NUMBER_COL + 1;
        public const int RC_NAME_COL = RC_NUMBER_COL + 2;
        public const int RC_MEASURE_COL = RC_NUMBER_COL + 3;
        public const int RC_QUANTITY_COL = RC_NUMBER_COL + 4;
        public const int RC_QUANTITY_FACT_COL = RC_NUMBER_COL + 5;
        public const int RC_LABOURNESS_COEFFICIENT_COL = RC_NUMBER_COL + 6;
        public const int RC_LABOURNESS_COL = RC_NUMBER_COL + 7;

        private Excel.Worksheet _worksheet;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public override Excel.Worksheet Worksheet
        {
            get { return base.Worksheet; }
            set
            {
            
                if (this.ReportCard != null) this.ReportCard.Worksheet = value;
                this.WorkersComposition.Worksheet = value;
                this.MachinesComposition.Worksheet = value;
               base.Worksheet = value;
            }
        }
        private string _code;

        public string Code
        {
            get { return _code; }
            set { SetProperty(ref _code, value); }
        }
        private decimal _labournessCoefficient;

        public decimal LabournessCoefficient
        {
            get { return _labournessCoefficient; }
            set { SetProperty(ref _labournessCoefficient, value); }
        }

        public override void UpdateExcelRepresetation()
        {
            RCWork rc_work = this;
            this.UpdateExellBindableObject();
            var rc_card = rc_work.ReportCard;
            if (rc_card != null)
            {
                rc_card.UpdateExcelRepresetation();
            }
        }

        public override int AdjustExcelRepresentionTree(int row)
        {
            RCWork rc_work = this;
            int rc_row = row;
            rc_work.ChangeTopRow(rc_row);
            ///Находимо работы с таким же номером и помещаем их ниже 
            int rc_work_cuont = 0;
    
            if (rc_work.ReportCard != null)
            {
                int rc_card_count = 0;
                rc_work.ReportCard.AdjustExcelRepresentionTree(rc_row);
               
                if (rc_work_cuont > rc_card_count)
                    rc_row += rc_work_cuont;
                else
                    rc_row += rc_card_count;
            }
            return rc_row;
        }


        public override void SetStyleFormats(int col)
        {
            RCWork rc_work = this;
            int ks_work_col = col;
            var rc_work_range = rc_work.GetRange(RC_LABOURNESS_COL);
            //     rc_work_range.SetBordersLine();
            rc_work_range.Interior.ColorIndex = ks_work_col;
            if (rc_work.ReportCard != null)
                rc_work.ReportCard.SetStyleFormats(ks_work_col);
        }

        public override object Clone()
        {
            RCWork new_work = (RCWork)base.Clone();
            new_work.Code = Code;
            new_work.LabournessCoefficient = LabournessCoefficient;
            return new_work;
        }


    }
}
