using ExellAddInsLib.MSG;

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

namespace MSGAddIn
{
    public partial class Ribbon1
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

       
        MSGExellModel MSGExellModel = new MSGExellModel();

        public ObservableCollection<MSGWork> MSGWorks { get; private set; } = new ObservableCollection<MSGWork>();
        public ObservableCollection<VOVRWork> VOVRWorks { get; private set; } = new ObservableCollection<VOVRWork>();
        public ObservableCollection<KSWork> KSWorks { get; private set; } = new  ObservableCollection<KSWork>();

        private int null_str_count = 0;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
   
        }

        private void buttonMSGLoad_Click(object sender, RibbonControlEventArgs e)
        {
            MSGExellModel.RegisterSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1];

            MSGExellModel.MSGWorks.Clear();
            VOVRWorks.Clear();
            KSWorks.Clear();

            LoadMSGWorks();
            LoadVOVRWorks();
            //LoadKSWorks();
            //LoadWorksReportCards();
            //CalcLabourness();
        }

        private void LoadMSGWorks()
        {
            Excel.Worksheet registerSheet = MSGExellModel.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            MSGExellModel.MSGWorks.Clear();

            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, MSG_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    MSGWork msg_work = new MSGWork();

                    msg_work.Number = registerSheet.Cells[rowIndex, MSG_NUMBER_COL].Value.ToString();
                    msg_work.CellAddressesMap.Add("Number", Tuple.Create(rowIndex, MSG_NUMBER_COL));

                    msg_work.Name = registerSheet.Cells[rowIndex, MSG_NAME_COL].Value;
                    msg_work.CellAddressesMap.Add("Name", Tuple.Create(rowIndex, MSG_NAME_COL));

                    if (registerSheet.Cells[rowIndex, MSG_MEASURE_COL].Value != null)
                    {
                        msg_work.UnitOfMeasurement = new UnitOfMeasurement(registerSheet.Cells[rowIndex, MSG_MEASURE_COL].Value);
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_MEASURE_COL], registerSheet.Cells[rowIndex, MSG_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                        msg_work.CellAddressesMap.Add("UnitOfMeasurement", Tuple.Create(rowIndex, MSG_MEASURE_COL));
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_MEASURE_COL], registerSheet.Cells[rowIndex, MSG_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, MSG_QUANTITY_COL].Value != null)
                    {
                        msg_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, MSG_QUANTITY_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_QUANTITY_COL], registerSheet.Cells[rowIndex, MSG_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                        msg_work.CellAddressesMap.Add("ProjectQuantity", Tuple.Create(rowIndex, MSG_QUANTITY_COL));

                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_QUANTITY_COL], registerSheet.Cells[rowIndex, MSG_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;
                 
                    if (registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value != null)
                    {
                        msg_work.Laboriousness = Decimal.Parse(registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL], registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                        msg_work.CellAddressesMap.Add("Laboriousness", Tuple.Create(rowIndex, MSG_LABOURNESS_COL));
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL], registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    DateTime start_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                    DateTime end_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());
                    WorkScheduleChunk work_sh_chunk = new WorkScheduleChunk(start_time, end_time);
                    work_sh_chunk.CellAddressesMap.Add("StartTime", Tuple.Create(rowIndex, MSG_START_DATE_COL));
                    work_sh_chunk.CellAddressesMap.Add("EndTime", Tuple.Create(rowIndex, MSG_END_DATE_COL));
                    msg_work.WorkSchedules.Add(work_sh_chunk);
                    MSGExellModel.Register(work_sh_chunk);

                    while (registerSheet.Cells[rowIndex + 1, MSG_NUMBER_COL].Value == null
                                 && registerSheet.Cells[rowIndex + 1, MSG_START_DATE_COL].Value != null)
                    {
                        rowIndex++;
                        start_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                        end_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());
                        WorkScheduleChunk  extra_work_sh_chunk = new WorkScheduleChunk(start_time, end_time);
                        extra_work_sh_chunk.CellAddressesMap.Add("StartTime", Tuple.Create(rowIndex, MSG_START_DATE_COL));
                        extra_work_sh_chunk.CellAddressesMap.Add("EndTime", Tuple.Create(rowIndex, MSG_END_DATE_COL));

                        msg_work.WorkSchedules.Add(extra_work_sh_chunk);
                        MSGExellModel.Register(extra_work_sh_chunk);
                    }
                    MSGExellModel.Register(msg_work);
                }
                rowIndex++;
            }
        }
        private void LoadVOVRWorks()
        {
            Excel.Worksheet registerSheet = MSGExellModel.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;


            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, VOVR_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    VOVRWork vovr_work = new VOVRWork();

                    vovr_work.Number = registerSheet.Cells[rowIndex, VOVR_NUMBER_COL].Value.ToString();
                    vovr_work.CellAddressesMap.Add("Number", Tuple.Create(rowIndex, VOVR_NUMBER_COL));

                    vovr_work.Name = registerSheet.Cells[rowIndex, VOVR_NAME_COL].Value.ToString();
                    vovr_work.CellAddressesMap.Add("Name", Tuple.Create(rowIndex, VOVR_NAME_COL));

                    if (registerSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value != null)
                    {
                        vovr_work.UnitOfMeasurement = new UnitOfMeasurement(registerSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value);
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_MEASURE_COL], registerSheet.Cells[rowIndex, VOVR_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                        vovr_work.CellAddressesMap.Add("UnitOfMeasurement", Tuple.Create(rowIndex, VOVR_MEASURE_COL));

                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_MEASURE_COL], registerSheet.Cells[rowIndex, VOVR_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL].Value != null)
                    {
                        vovr_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL], registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                        vovr_work.CellAddressesMap.Add("ProjectQuantity", Tuple.Create(rowIndex, VOVR_QUANTITY_COL));

                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL], registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;
                  
                    if (registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL].Value != null)
                    {
                        vovr_work.Laboriousness = Decimal.Parse(registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL], registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                        vovr_work.CellAddressesMap.Add("Laboriousness", Tuple.Create(rowIndex, VOVR_LABOURNESS_COL));

                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL], registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    MSGExellModel.Register(vovr_work);
                   
                }

                rowIndex++;
            }
        }
        private void LoadKSWorks()
        {
            Excel.Worksheet registerSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1];
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, KS_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    KSWork ks_work = new  KSWork();

                    ks_work.Number = registerSheet.Cells[rowIndex, KS_NUMBER_COL].Value.ToString();
                    ks_work.Name = registerSheet.Cells[rowIndex, KS_NAME_COL].Value;
                    if (registerSheet.Cells[rowIndex, KS_MEASURE_COL].Value != null)
                    {
                        ks_work.UnitOfMeasurement = new UnitOfMeasurement(registerSheet.Cells[rowIndex, KS_MEASURE_COL].Value);
                        registerSheet.Range[registerSheet.Cells[rowIndex, KS_MEASURE_COL], registerSheet.Cells[rowIndex, KS_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, KS_MEASURE_COL], registerSheet.Cells[rowIndex, KS_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, KS_QUANTITY_COL].Value != null)
                    {
                        ks_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, KS_QUANTITY_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, KS_QUANTITY_COL], registerSheet.Cells[rowIndex, KS_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, KS_QUANTITY_COL], registerSheet.Cells[rowIndex, KS_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;


                    if (registerSheet.Cells[rowIndex, KS_LABOURNESS_COL].Value != null)
                    {
                        ks_work.Laboriousness = Decimal.Parse(registerSheet.Cells[rowIndex, KS_LABOURNESS_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, KS_LABOURNESS_COL], registerSheet.Cells[rowIndex, KS_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, KS_LABOURNESS_COL], registerSheet.Cells[rowIndex, KS_LABOURNESS_COL]].Interior.Color
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
        private void LoadWorksReportCards()
        {
            Excel.Worksheet registerSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1];
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, WRC_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    WorkReportCard report_card = new  WorkReportCard();
                    DateTime end_date = DateTime.Parse(registerSheet.Cells[WORKS_END_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
                    report_card.Number = registerSheet.Cells[rowIndex, WRC_NUMBER_COL].Value.ToString();
                    int date_index = 0;
                    while (DateTime.Parse(registerSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL+ date_index].Value.ToString()) < end_date)
                    {
                        DateTime current_date = DateTime.Parse(registerSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString());
                        decimal quantity =0;
                        if(registerSheet.Cells[rowIndex, WRC_DATE_COL + date_index].Value!=null)
                            quantity = Decimal.Parse(registerSheet.Cells[rowIndex, WRC_DATE_COL + date_index].Value.ToString());
                        if(quantity!=0)
                        {
                            WorkDay workDay = new WorkDay();
                            workDay.Date = current_date;
                            workDay.Quantity = quantity;
                            report_card.Add(workDay);
                        }
                        date_index++;
                    }
                    VOVRWork vovr_work = VOVRWorks.Where(w => w.Number== report_card.Number).FirstOrDefault();
                    if (vovr_work != null && report_card.Count>0)
                        vovr_work.ReportCard= report_card;
                }
                rowIndex++;
            }

        }
        private void CalcLabourness()
        {
            foreach(MSGWork msg_work in MSGWorks)
            {
                if(msg_work.Laboriousness==0)
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
                       common_vovr_laboueness +=vovr_work.ProjectQuantity*vovr_work.Laboriousness;
                    }
                    msg_work.Laboriousness = common_vovr_laboueness / msg_work.ProjectQuantity;
                }
            }
        }

        private void btnNotifyTest_Click(object sender, RibbonControlEventArgs e)
        {
            MSGExellModel.VOVRWorks[0].Name = "2222222";

        }
    }
}
