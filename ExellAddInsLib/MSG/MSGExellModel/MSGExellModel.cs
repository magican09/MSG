using System;
using System.Collections.ObjectModel;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.ComponentModel;

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

        private int null_str_count = 0;
        public ObservableCollection<MSGWork> MSGWorks { get; private set; } = new ObservableCollection<MSGWork>();
        public ObservableCollection<VOVRWork> VOVRWorks { get; private set; } = new ObservableCollection<VOVRWork>();
        public ObservableCollection<KSWork> KSWorks { get; private set; } = new ObservableCollection<KSWork>();
        public ObservableCollection<UnitOfMeasurement> UnitOfMeasurements { get; set; } = new ObservableCollection<UnitOfMeasurement>();
        public MSGExellModel Owner { get; set; }
        public Excel.Worksheet RegisterSheet { get; set; }

        public MSGExellModel()
        {

        }

        public void Register(object work)
        {
            if (work is INotifyPropertyChanged notified_object)
                notified_object.PropertyChanged += OnPropertyChange;
            switch (work.GetType().Name)
            {
                case nameof(MSGWork):
                    {

                        MSGWork msg_work = (MSGWork)work;
                        if (!MSGWorks.Contains(msg_work))
                            MSGWorks.Add(msg_work);
                        break;
                    }
                case nameof(VOVRWork):
                    {
                        VOVRWork vovr_work = (VOVRWork)work;
                        if (!this.VOVRWorks.Contains(vovr_work))
                            this.VOVRWorks.Add(vovr_work);

                        MSGWork msg_work = this.MSGWorks.Where(w => w.Number.StartsWith(vovr_work.Number.Remove(vovr_work.Number.LastIndexOf(".")))).FirstOrDefault();
                        if (msg_work != null)
                        {
                            msg_work.VOVRWorks.Add(vovr_work);
                        }

                        break;
                    }
                case nameof(KSWork):
                    {
                        KSWork ks_work = (KSWork)work;
                        if (!this.KSWorks.Contains(ks_work))
                            this.KSWorks.Add(ks_work);

                        VOVRWork vovr_work = VOVRWorks.Where(w => w.Number.StartsWith(ks_work.Number.Remove(ks_work.Number.LastIndexOf(".")))).FirstOrDefault();
                        KSWorks.Add(ks_work);
                        if (vovr_work != null)
                            vovr_work.KSWorks.Add(ks_work);

                        break;
                    }

                case nameof(WorkReportCard):
                    {
                        WorkReportCard report_card = (WorkReportCard)work;

                        KSWork ks_work = KSWorks.Where(w => w.Number == report_card.Number).FirstOrDefault();
                        if (ks_work != null && report_card.Count > 0)
                            ks_work.ReportCard = report_card;

                        break;
                    }

            }
        }
        private void OnPropertyChange(object sender, PropertyChangedEventArgs e)
        {
            if (sender is IExcelBindableBase bindable_object)
            {
                if (bindable_object.CellAddressesMap.ContainsKey(e.PropertyName))
                {
                    RegisterSheet.Cells[bindable_object.CellAddressesMap[e.PropertyName].Item1,
                   bindable_object.CellAddressesMap[e.PropertyName].Item2] = sender.GetType().GetProperty(e.PropertyName).GetValue(sender).ToString();
                }
            }
        }

        public void ResetMSG_VOVR()
        {
            foreach (MSGWork msg_work in this.MSGWorks)
            {
                msg_work.Quantity = 0;
                msg_work.Laboriousness = 0;
            }
            foreach (VOVRWork vovr_work in this.VOVRWorks)
            {
                vovr_work.Quantity = 0;
                vovr_work.Laboriousness = 0;
            }
        }
        public void LoadMSGWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            this.MSGWorks.Clear();

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

                    //if (registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value!=null )
                    //     msg_work.Laboriousness =Decimal.Parse(registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value.ToString());

                    //  msg_work.CellAddressesMap.Add("Laboriousness", Tuple.Create(rowIndex, MSG_LABOURNESS_COL));
                    msg_work.CellAddressesMap.Add("UnitOfMeasurement", Tuple.Create(rowIndex, MSG_MEASURE_COL));
                    msg_work.CellAddressesMap.Add("ProjectQuantity", Tuple.Create(rowIndex, MSG_QUANTITY_COL));
                    msg_work.CellAddressesMap.Add("Quantity", Tuple.Create(rowIndex, MSG_QUANTITY_FACT_COL));
                    msg_work.CellAddressesMap.Add("Laboriousness", Tuple.Create(rowIndex, MSG_LABOURNESS_COL));

                    if (registerSheet.Cells[rowIndex, MSG_MEASURE_COL].Value != null)
                    {
                        string un_name = registerSheet.Cells[rowIndex, MSG_MEASURE_COL].Value.ToString();
                        UnitOfMeasurement unitOfMeasurement = UnitOfMeasurements.FirstOrDefault(um => um.Name == un_name);
                        if (unitOfMeasurement != null)
                        {
                            msg_work.UnitOfMeasurement = unitOfMeasurement;
                            registerSheet.Range[registerSheet.Cells[rowIndex, MSG_MEASURE_COL], registerSheet.Cells[rowIndex, MSG_MEASURE_COL]].Interior.Color
                                = XlRgbColor.rgbWhite;
                        }
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_MEASURE_COL], registerSheet.Cells[rowIndex, MSG_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, MSG_QUANTITY_COL].Value != null)
                    {
                        msg_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, MSG_QUANTITY_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_QUANTITY_COL], registerSheet.Cells[rowIndex, MSG_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;

                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_QUANTITY_COL], registerSheet.Cells[rowIndex, MSG_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value != null)
                    {
                        msg_work.Laboriousness = Decimal.Parse(registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL], registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
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
                    this.Register(work_sh_chunk);

                    while (registerSheet.Cells[rowIndex + 1, MSG_NUMBER_COL].Value == null
                                 && registerSheet.Cells[rowIndex + 1, MSG_START_DATE_COL].Value != null)
                    {
                        rowIndex++;
                        start_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                        end_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());
                        WorkScheduleChunk extra_work_sh_chunk = new WorkScheduleChunk(start_time, end_time);
                        extra_work_sh_chunk.CellAddressesMap.Add("StartTime", Tuple.Create(rowIndex, MSG_START_DATE_COL));
                        extra_work_sh_chunk.CellAddressesMap.Add("EndTime", Tuple.Create(rowIndex, MSG_END_DATE_COL));

                        msg_work.WorkSchedules.Add(extra_work_sh_chunk);
                        this.Register(extra_work_sh_chunk);
                    }
                    this.Register(msg_work);
                }
                rowIndex++;
            }
        }
        public void LoadVOVRWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
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
                    vovr_work.CellAddressesMap.Add("UnitOfMeasurement", Tuple.Create(rowIndex, VOVR_MEASURE_COL));
                    vovr_work.CellAddressesMap.Add("ProjectQuantity", Tuple.Create(rowIndex, VOVR_QUANTITY_COL));
                    vovr_work.CellAddressesMap.Add("Quantity", Tuple.Create(rowIndex, VOVR_QUANTITY_FACT_COL));
                    vovr_work.CellAddressesMap.Add("Laboriousness", Tuple.Create(rowIndex, VOVR_LABOURNESS_COL));


                    if (registerSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value != null)
                    {
                        vovr_work.UnitOfMeasurement = new UnitOfMeasurement(registerSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value);
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_MEASURE_COL], registerSheet.Cells[rowIndex, VOVR_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;

                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_MEASURE_COL], registerSheet.Cells[rowIndex, VOVR_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL].Value != null)
                    {
                        vovr_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL], registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;

                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL], registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL].Value != null)
                    {
                        vovr_work.Laboriousness = Decimal.Parse(registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL], registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL], registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbRed;


                    this.Register(vovr_work);

                }

                rowIndex++;
            }
        }
        public void LoadKSWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, KS_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    KSWork ks_work = new KSWork();

                    ks_work.Number = registerSheet.Cells[rowIndex, KS_NUMBER_COL].Value.ToString();
                    ks_work.CellAddressesMap.Add("Number", Tuple.Create(rowIndex, KS_NUMBER_COL));

                    ks_work.Name = registerSheet.Cells[rowIndex, KS_NAME_COL].Value;
                    ks_work.CellAddressesMap.Add("Name", Tuple.Create(rowIndex, KS_NAME_COL));

                    ks_work.CellAddressesMap.Add("UnitOfMeasurement", Tuple.Create(rowIndex, KS_MEASURE_COL));
                    ks_work.CellAddressesMap.Add("ProjectQuantity", Tuple.Create(rowIndex, KS_QUANTITY_COL));
                    ks_work.CellAddressesMap.Add("Quantity", Tuple.Create(rowIndex, KS_QUANTITY_FACT_COL));
                    ks_work.CellAddressesMap.Add("Laboriousness", Tuple.Create(rowIndex, KS_LABOURNESS_COL));

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

                    this.Register(ks_work);
                }
                rowIndex++;
            }
        }
        public void LoadWorksReportCards()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, WRC_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    WorkReportCard report_card = new WorkReportCard();
                    DateTime end_date = DateTime.Parse(registerSheet.Cells[WORKS_END_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());

                    report_card.Number = registerSheet.Cells[rowIndex, WRC_NUMBER_COL].Value.ToString();
                    report_card.CellAddressesMap.Add("Number", Tuple.Create(rowIndex, WRC_NUMBER_COL));

                    int date_index = 0;
                    while (DateTime.Parse(registerSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString()) < end_date)
                    {
                        DateTime current_date = DateTime.Parse(registerSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString());
                        decimal quantity = 0;
                        if (registerSheet.Cells[rowIndex, WRC_DATE_COL + date_index].Value != null)
                            quantity = Decimal.Parse(registerSheet.Cells[rowIndex, WRC_DATE_COL + date_index].Value.ToString());
                        if (quantity != 0)
                        {
                            WorkDay workDay = new WorkDay();
                            workDay.Date = current_date;
                            workDay.CellAddressesMap.Add("Date", Tuple.Create(WRC_DATE_ROW, WRC_DATE_COL + date_index));
                            workDay.Quantity = quantity;
                            workDay.CellAddressesMap.Add("Quantity", Tuple.Create(rowIndex, WRC_DATE_COL + date_index));
                            this.Register(workDay);
                            report_card.Add(workDay);
                        }
                        date_index++;
                    }
                    this.Register(report_card);
                }
                rowIndex++;
            }

        }
        public void CalcLabourness()
        {
            foreach (MSGWork msg_work in this.MSGWorks)
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
        public void CalcQuantity()
        {
            foreach (MSGWork msg_work in this.MSGWorks)
            {
                if (msg_work.Quantity == 0)
                {
                    decimal common_vovr_quantity = 0;
                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        if (vovr_work.Quantity == 0)
                        {
                            decimal common_ks_labour_quantity = 0;

                            foreach (KSWork ks_work in vovr_work.KSWorks)
                            {
                                if (ks_work.Laboriousness != 0)
                                {
                                    ks_work.Quantity = 0;
                                    foreach (WorkDay day in ks_work.ReportCard)
                                    {
                                        ks_work.Quantity += day.Quantity;
                                    }
                                    decimal ks_labour_quantity = ks_work.Quantity / ks_work.Laboriousness;
                                    common_ks_labour_quantity += ks_labour_quantity;
                                }

                            }

                            vovr_work.Quantity = common_ks_labour_quantity;

                        }
                        common_vovr_quantity += vovr_work.Quantity;
                    }
                    msg_work.Quantity = common_vovr_quantity;
                }
            }
        }
        public void RealoadAll()
        {
            this.MSGWorks.Clear();
            this.VOVRWorks.Clear();
            this.KSWorks.Clear();

            this.LoadMSGWorks();
            this.LoadVOVRWorks();
            this.LoadKSWorks();
            this.LoadWorksReportCards();
        }
    }
}
