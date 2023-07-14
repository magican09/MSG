using System;
using System.Collections.ObjectModel;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace ExellAddInsLib.MSG
{
    public class MSGExellModel
    {
        public const int WORKS_END_DATE_ROW = 2;
        public const int WORKS_END_DATE_COL = 3;

        public const int FIRST_ROW_INDEX = 7;
        public const int MSG_NUMBER_COL = 2;
        public const int MSG_NAME_COL = 3;
        public const int MSG_MEASURE_COL = 4;
        public const int MSG_QUANTITY_COL = 5;
        public const int MSG_QUANTITY_FACT_COL = 6;
        public const int MSG_LABOURNESS_COL = 7;
        public const int MSG_START_DATE_COL = 8;
        public const int MSG_END_DATE_COL = 9;


        public const int VOVR_NUMBER_COL = 10;
        public const int VOVR_NAME_COL = 11;
        public const int VOVR_MEASURE_COL = 12;
        public const int VOVR_QUANTITY_COL = 13;
        public const int VOVR_QUANTITY_FACT_COL = 14;
        public const int VOVR_LABOURNESS_COL = 15;


        public const int KS_NUMBER_COL = 16;
        public const int KS_CODE_COL = 17;
        public const int KS_NAME_COL = 18;
        public const int KS_MEASURE_COL = 19;
        public const int KS_QUANTITY_COL = 20;
        public const int KS_QUANTITY_FACT_COL = 21;
        public const int KS_LABOURNESS_COL = 22;

        public const int WRC_DATE_ROW = 6;
        public const int WRC_NUMBER_COL = 23;
        public const int WRC_DATE_COL = 24;

       
        public  ObservableCollection<MSGWork> MSGWorks { get; private set; } = new ObservableCollection<MSGWork>();
        public  ObservableCollection<VOVRWork> VOVRWorks { get; private set; } = new ObservableCollection<VOVRWork>();
        public  ObservableCollection<KSWork> KSWorks { get; private set; } = new ObservableCollection<KSWork>();

        
        public   Excel.Worksheet RegisterSheet { get; set; }
       
        public MSGExellModel()
        {

        }

        public void Register(Excel.Range range, string propName, object work)
        {
            switch(work.GetType().Name)
            {
                case nameof(MSGWork):
                    {
                        MSGWork msg_work = (MSGWork)work;

                        break;
                    }
            }
        }

        public   void LoadMSGWorks()
        {
            
            int rowIndex = FIRST_ROW_INDEX;
            int null_str_count = 0;
            MSGWorks.Clear();

            while (null_str_count < 100)
            {
                if (RegisterSheet.Cells[rowIndex, MSG_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    MSGWork msg_work = new MSGWork();

                    msg_work.Number = RegisterSheet.Cells[rowIndex, MSG_NUMBER_COL].Value.ToString();
                    msg_work.Name = RegisterSheet.Cells[rowIndex, MSG_NAME_COL].Value;
                    if (RegisterSheet.Cells[rowIndex, MSG_MEASURE_COL].Value != null)
                    {
                        msg_work.UnitOfMeasurement = new UnitOfMeasurement(RegisterSheet.Cells[rowIndex, MSG_MEASURE_COL].Value);
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, MSG_MEASURE_COL], RegisterSheet.Cells[rowIndex, MSG_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, MSG_MEASURE_COL], RegisterSheet.Cells[rowIndex, MSG_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (RegisterSheet.Cells[rowIndex, MSG_QUANTITY_COL].Value != null)
                    {
                        msg_work.ProjectQuantity = Decimal.Parse(RegisterSheet.Cells[rowIndex, MSG_QUANTITY_COL].Value.ToString());
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, MSG_QUANTITY_COL], RegisterSheet.Cells[rowIndex, MSG_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, MSG_QUANTITY_COL], RegisterSheet.Cells[rowIndex, MSG_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (RegisterSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value != null)
                    {
                        msg_work.Laboriousness = Decimal.Parse(RegisterSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value.ToString());
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, MSG_LABOURNESS_COL], RegisterSheet.Cells[rowIndex, MSG_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, MSG_LABOURNESS_COL], RegisterSheet.Cells[rowIndex, MSG_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    DateTime start_time = DateTime.Parse(RegisterSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                    DateTime end_time = DateTime.Parse(RegisterSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());
                    msg_work.WorkSchedules.Add(new WorkScheduleChunk(start_time, end_time));
                    while (RegisterSheet.Cells[rowIndex + 1, MSG_NUMBER_COL].Value == null
                                 && RegisterSheet.Cells[rowIndex + 1, MSG_START_DATE_COL].Value != null)
                    {
                        rowIndex++;
                        start_time = DateTime.Parse(RegisterSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                        end_time = DateTime.Parse(RegisterSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());
                        msg_work.WorkSchedules.Add(new WorkScheduleChunk(start_time, end_time));
                    }
                    MSGWorks.Add(msg_work);
                }
                rowIndex++;
            }
        }
        public   void LoadVOVRWorks()
        {
            
            int rowIndex = FIRST_ROW_INDEX;
            int null_str_count = 0;


            while (null_str_count < 100)
            {
                if (RegisterSheet.Cells[rowIndex, VOVR_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    VOVRWork vovr_work = new VOVRWork();

                    vovr_work.Number = RegisterSheet.Cells[rowIndex, VOVR_NUMBER_COL].Value.ToString();
                    vovr_work.Name = RegisterSheet.Cells[rowIndex, VOVR_NAME_COL].Value;
                    if (RegisterSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value != null)
                    {
                        vovr_work.UnitOfMeasurement = new UnitOfMeasurement(RegisterSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value);
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, VOVR_MEASURE_COL], RegisterSheet.Cells[rowIndex, VOVR_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, VOVR_MEASURE_COL], RegisterSheet.Cells[rowIndex, VOVR_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (RegisterSheet.Cells[rowIndex, VOVR_QUANTITY_COL].Value != null)
                    {
                        vovr_work.ProjectQuantity = Decimal.Parse(RegisterSheet.Cells[rowIndex, VOVR_QUANTITY_COL].Value.ToString());
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, VOVR_QUANTITY_COL], RegisterSheet.Cells[rowIndex, VOVR_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, VOVR_QUANTITY_COL], RegisterSheet.Cells[rowIndex, VOVR_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (RegisterSheet.Cells[rowIndex, VOVR_LABOURNESS_COL].Value != null)
                    {
                        vovr_work.Laboriousness = Decimal.Parse(RegisterSheet.Cells[rowIndex, VOVR_LABOURNESS_COL].Value.ToString());
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, VOVR_LABOURNESS_COL], RegisterSheet.Cells[rowIndex, VOVR_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, VOVR_LABOURNESS_COL], RegisterSheet.Cells[rowIndex, VOVR_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    VOVRWorks.Add(vovr_work);
                    MSGWork msg_work = MSGWorks.Where(w => w.Number.StartsWith(vovr_work.Number.Remove(vovr_work.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (msg_work != null)
                    {
                        msg_work.VOVRWorks.Add(vovr_work);
                    }
                }
                rowIndex++;
            }
        }
        public   void LoadKSWorks()
        {
            
            int rowIndex = FIRST_ROW_INDEX;
           int  null_str_count = 0;
            while (null_str_count < 100)
            {
                if (RegisterSheet.Cells[rowIndex, KS_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    KSWork ks_work = new KSWork();

                    ks_work.Number = RegisterSheet.Cells[rowIndex, KS_NUMBER_COL].Value.ToString();
                    ks_work.Name = RegisterSheet.Cells[rowIndex, KS_NAME_COL].Value;
                    if (RegisterSheet.Cells[rowIndex, KS_MEASURE_COL].Value != null)
                    {
                        ks_work.UnitOfMeasurement = new UnitOfMeasurement(RegisterSheet.Cells[rowIndex, KS_MEASURE_COL].Value);
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, KS_MEASURE_COL], RegisterSheet.Cells[rowIndex, KS_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, KS_MEASURE_COL], RegisterSheet.Cells[rowIndex, KS_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (RegisterSheet.Cells[rowIndex, KS_QUANTITY_COL].Value != null)
                    {
                        ks_work.ProjectQuantity = Decimal.Parse(RegisterSheet.Cells[rowIndex, KS_QUANTITY_COL].Value.ToString());
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, KS_QUANTITY_COL], RegisterSheet.Cells[rowIndex, KS_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, KS_QUANTITY_COL], RegisterSheet.Cells[rowIndex, KS_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;


                    if (RegisterSheet.Cells[rowIndex, KS_LABOURNESS_COL].Value != null)
                    {
                        ks_work.Laboriousness = Decimal.Parse(RegisterSheet.Cells[rowIndex, KS_LABOURNESS_COL].Value.ToString());
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, KS_LABOURNESS_COL], RegisterSheet.Cells[rowIndex, KS_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        RegisterSheet.Range[RegisterSheet.Cells[rowIndex, KS_LABOURNESS_COL], RegisterSheet.Cells[rowIndex, KS_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    var st = ks_work.Number.Remove(ks_work.Number.LastIndexOf("."));
                    VOVRWork vovr_work = VOVRWorks.Where(w => w.Number.StartsWith(ks_work.Number.Remove(ks_work.Number.LastIndexOf(".")))).FirstOrDefault();
                    KSWorks.Add(ks_work);
                    if (vovr_work != null)
                        vovr_work.KSWorks.Add(ks_work);
                }
                rowIndex++;
            }
        }
        public   void LoadWorksReportCards()
        {
            
            int rowIndex = FIRST_ROW_INDEX;
            int null_str_count = 0;
            while (null_str_count < 100)
            {
                if (RegisterSheet.Cells[rowIndex, WRC_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    WorkReportCard report_card = new WorkReportCard();
                    DateTime end_date = DateTime.Parse(RegisterSheet.Cells[WORKS_END_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
                    report_card.Number = RegisterSheet.Cells[rowIndex, WRC_NUMBER_COL].Value.ToString();
                    int date_index = 0;
                    while (DateTime.Parse(RegisterSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString()) < end_date)
                    {
                        DateTime current_date = DateTime.Parse(RegisterSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString());
                        decimal quantity = 0;
                        if (RegisterSheet.Cells[rowIndex, WRC_DATE_COL + date_index].Value != null)
                            quantity = Decimal.Parse(RegisterSheet.Cells[rowIndex, WRC_DATE_COL + date_index].Value.ToString());
                        if (quantity != 0)
                        {
                            WorkDay workDay = new WorkDay();
                            workDay.Date = current_date;
                            workDay.Quantity = quantity;
                            report_card.Add(workDay);
                        }
                        date_index++;
                    }
                    VOVRWork vovr_work = VOVRWorks.Where(w => w.Number == report_card.Number).FirstOrDefault();
                    if (vovr_work != null && report_card.Count > 0)
                        vovr_work.ReportCard = report_card;
                }
                rowIndex++;
            }

        }
        public  void CalcLabourness()
        {
            foreach (MSGWork msg_work in MSGWorks)
            {
                if (msg_work.Laboriousness == 0)
                {
                    decimal common_vovr_laboueness = 0;
                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        if (vovr_work.Laboriousness == 0)
                        {
                            decimal common_ks_laboueness = 0;
                            foreach (KSWork ks_work in vovr_work.KSWorks)
                            {
                                common_ks_laboueness += ks_work.ProjectQuantity * ks_work.Laboriousness;
                            }
                            vovr_work.Laboriousness = common_ks_laboueness / vovr_work.ProjectQuantity;
                        }
                        common_vovr_laboueness += vovr_work.ProjectQuantity * vovr_work.Laboriousness;
                    }
                    msg_work.Laboriousness = common_vovr_laboueness / msg_work.ProjectQuantity;
                }
            }
        }

    }
}
