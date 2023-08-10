using ExellAddInsLib.MSG.Section;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using static System.Collections.Specialized.BitVector32;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public  partial class MSGExellModel : ExellModelBase
    {
        /// <summary>
        /// Констраныт номеров строк и стобцов в документе exel 
        /// </summary>

        ///Начальная
        public const int COMMON_PARAMETRS_VALUE_COL = 3; //Номер стобца с общим параметрами проекта

        public const int CONTRACT_CODE_ROW = 2; //Код объекта или договора
        public const int CONSTRUCTION_OBJECT_CODE_ROW = 3;// Код объекта
        public const int CONSTRUCTION_SUBOBJECT_CODE_ROW = 4;//Код подъобьекта

        /// <summary>
        /// Ведомость_
        /// </summary>
        public const int WORKS_START_DATE_ROW = 1;
        public const int WORKS_START_DATE_COL = 3;
        public const int WORKS_END_DATE_ROW = 2;
        public const int WORKS_END_DATE_COL = 3;

        public const int FIRST_ROW_INDEX = 8;

        public const int WSEC_NUMBER_COL = 2;
        public const int WSEC_NAME_COL = WSEC_NUMBER_COL + 1;

        public const int MSG_NUMBER_COL = 4;
        public const int MSG_NAME_COL = MSG_NUMBER_COL + 1;
        public const int MSG_MEASURE_COL = MSG_NUMBER_COL + 2;
        public const int MSG_QUANTITY_COL = MSG_NUMBER_COL + 3;
        public const int MSG_QUANTITY_FACT_COL = MSG_NUMBER_COL + 4;
        public const int MSG_LABOURNESS_COL = MSG_NUMBER_COL + 5;
        public const int MSG_START_DATE_COL = MSG_NUMBER_COL + 6;
        public const int MSG_END_DATE_COL = MSG_NUMBER_COL + 7;
        public const int MSG_SUNDAY_IS_VOCATION_COL = MSG_NUMBER_COL + 8;

        public const int MSG_NEEDS_OF_WORKERS_NUMBER_COL = MSG_SUNDAY_IS_VOCATION_COL + 1;
        public const int MSG_NEEDS_OF_WORKERS_NAME_COL = MSG_NEEDS_OF_WORKERS_NUMBER_COL + 1;
        public const int MSG_NEEDS_OF_WORKERS_QUANTITY_COL = MSG_NEEDS_OF_WORKERS_NUMBER_COL + 2;

        public const int MSG_NEEDS_OF_MACHINE_NUMBER_COL = MSG_NEEDS_OF_WORKERS_QUANTITY_COL + 1;
        public const int MSG_NEEDS_OF_MACHINE_NAME_COL = MSG_NEEDS_OF_MACHINE_NUMBER_COL + 1;
        public const int MSG_NEEDS_OF_MACHINE_QUANTITY_COL = MSG_NEEDS_OF_MACHINE_NUMBER_COL + 2;


        public const int VOVR_NUMBER_COL = MSG_NEEDS_OF_MACHINE_QUANTITY_COL + 1;
        public const int VOVR_NAME_COL = VOVR_NUMBER_COL + 1;
        public const int VOVR_MEASURE_COL = VOVR_NUMBER_COL + 2;
        public const int VOVR_QUANTITY_COL = VOVR_NUMBER_COL + 3;
        public const int VOVR_QUANTITY_FACT_COL = VOVR_NUMBER_COL + 4;
        public const int VOVR_LABOURNESS_COL = VOVR_NUMBER_COL + 5;


        public const int KS_NUMBER_COL = VOVR_LABOURNESS_COL + 1;
        public const int KS_CODE_COL = KS_NUMBER_COL + 1;
        public const int KS_NAME_COL = KS_NUMBER_COL + 2;
        public const int KS_MEASURE_COL = KS_NUMBER_COL + 3;
        public const int KS_QUANTITY_COL = KS_NUMBER_COL + 4;
        public const int KS_QUANTITY_FACT_COL = KS_NUMBER_COL + 5;
        public const int KS_LABOURNESS_COL = KS_NUMBER_COL + 6;

        public const int RC_NUMBER_COL = KS_LABOURNESS_COL + 1;
        public const int RC_CODE_COL = RC_NUMBER_COL + 1;
        public const int RC_NAME_COL = RC_NUMBER_COL + 2;
        public const int RC_MEASURE_COL = RC_NUMBER_COL + 3;
        public const int RC_QUANTITY_COL = RC_NUMBER_COL + 4;
        public const int RC_QUANTITY_FACT_COL = RC_NUMBER_COL + 5;
        public const int RC_LABOURNESS_COEFFICIENT_COL = RC_NUMBER_COL + 6;
        public const int RC_LABOURNESS_COL = RC_NUMBER_COL + 7;


        public const int WRC_DATE_ROW = 6;

        public const int WRC_NUMBER_COL = RC_LABOURNESS_COL + 1;
        public const int WRC_PC_QUANTITY_COL = WRC_NUMBER_COL + 1;
        public const int WRC_DATE_COL = WRC_PC_QUANTITY_COL + 1;

        /// <summary>
        /// Люди_
        /// </summary>
        public const int W_CONSUMPTIONS_FIRST_ROW_INDEX = 4;
        public const int W_CONSUMPTIONS_NUMBER_COL = 1;
        public const int W_CONSUMPTIONS_NAME_COL = 2;
        public const int W_CONSUMPTIONS_DATE_RAW = 3;
        public const int W_CONSUMPTIONS_FIRST_DATE_COL = 3;

        public const int MCH_CONSUMPTIONS_FIRST_ROW_INDEX = 4;
        public const int MCH_CONSUMPTIONS_NUMBER_COL = 1;
        public const int MCH_CONSUMPTIONS_NAME_COL = 2;
        public const int MCH_CONSUMPTIONS_DATE_RAW = 3;
        public const int MCH_CONSUMPTIONS_FIRST_DATE_COL = 3;

        public const int _SECTIONS_GAP = 2;
        
        public const int W_SECTION_COLOR = 33;


        private int null_str_count = 0;
        /// <summary>
        /// Дата начала ведомости 
        /// </summary>
        public DateTime WorksStartDate { get; set; }
        /// <summary>
        /// Дата окончания работ в данной ведомости в соотвествии с планируемыми в  сроками отраженнным в части МСГ ведомости.
        /// (в части WorkSchedules работ MSGWork)
        /// </summary>
        public DateTime WorksEndDate
        {
            get
            {
                DateTime end_date = DateTime.MinValue;
                var last_ended_work = this.MSGWorks.OrderBy(w => w.WorkSchedules.EndDate).LastOrDefault();
                if (last_ended_work != null)
                    end_date = last_ended_work.WorkSchedules.EndDate;
                return end_date;
            }

        }


        /// <summary>
        /// Общее количество каленадрых дне с начала до окончания работ
        /// </summary>
        public int WorkedDaysNumber
        {
            get
            {
                return (this.WorksEndDate - this.WorksStartDate).Days;
            }

        }



        private ExcelNotifyChangedCollection<WorksSection> _worksSections;
        /// <summary>
        /// Коллекция с разделами работ
        /// </summary>
        public ExcelNotifyChangedCollection<WorksSection> WorksSections
        {
            get { return _worksSections; }
            set { SetProperty(ref _worksSections, value); }
        }


        private ExcelNotifyChangedCollection<MSGWork> _mSGWorks;
        /// <summary>
        /// Коллекция с работами типа МСГ модели
        /// </summary>
        public ExcelNotifyChangedCollection<MSGWork> MSGWorks
        {
            get { return _mSGWorks; }
            set { SetProperty(ref _mSGWorks, value); }
        }

        private ExcelNotifyChangedCollection<VOVRWork> _vOVRWorks;
        /// <summary>
        /// Коллекция с работами типа ВОВР модели
        /// </summary>
        public ExcelNotifyChangedCollection<VOVRWork> VOVRWorks
        {
            get { return _vOVRWorks; }
            set { SetProperty(ref _vOVRWorks, value); }
        }


        private ExcelNotifyChangedCollection<KSWork> _kSWorks;
        /// <summary>
        /// Коллекция с работами типа КС-2 модели
        /// </summary>
        public ExcelNotifyChangedCollection<KSWork> KSWorks
        {
            get { return _kSWorks; }
            set { SetProperty(ref _kSWorks, value); }
        }


        private ExcelNotifyChangedCollection<RCWork> _rCWorks;
        /// <summary>
        /// Коллекция с работами типа ждя учета модели
        /// </summary>
        public ExcelNotifyChangedCollection<RCWork> RCWorks
        {
            get { return _rCWorks; }
            set { SetProperty(ref _rCWorks, value); }
        }

        private ExcelNotifyChangedCollection<WorkReportCard> _workReportCards;
        /// <summary>
        /// Коллекция  табелей выполненных работ
        /// </summary>
        public ExcelNotifyChangedCollection<WorkReportCard> WorkReportCards
        {
            get { return _workReportCards; }
            set { SetProperty(ref _workReportCards, value); }
        }

        private ObservableCollection<IExcelBindableBase> _invalidObjects = new ObservableCollection<IExcelBindableBase>();


        private ExcelNotifyChangedCollection<UnitOfMeasurement> _unitOfMeasurements;
        /// <summary>
        /// Коллекция с единицами измерения модели
        /// </summary>
        public ExcelNotifyChangedCollection<UnitOfMeasurement> UnitOfMeasurements
        {
            get { return _unitOfMeasurements; }
            set { SetProperty(ref _unitOfMeasurements, value); }
        }

        private WorkersComposition _WorkersComposition;
        /// <summary>
        /// Состав работников ( потребности)
        /// </summary>
        public WorkersComposition WorkersComposition
        {
            get { return _WorkersComposition; }
            set { SetProperty(ref _WorkersComposition, value); }
        }
        private MachinesComposition _machinesComposition;
        /// <summary>
        /// Состав работников ( потребности)
        /// </summary>
        public MachinesComposition MachinesComposition
        {
            get { return _machinesComposition; }
            set { SetProperty(ref _machinesComposition, value); }
        }

        private WorkerConsumptions _workerConsumptions;
        /// <summary>
        /// Потребления работников
        /// </summary>
        public WorkerConsumptions WorkerConsumptions
        {
            get { return _workerConsumptions; }
            set { SetProperty(ref _workerConsumptions, value); }
        }

        private MachineConsumptions _machineConsumptions;
        public MachineConsumptions MachineConsumptions
        {
            get { return _machineConsumptions; }
            set { SetProperty(ref _machineConsumptions, value); }
        }
        /// <summary>
        ///Шифр объекта или договора
        /// </summary>
        public string ContractCode { get; set; }
        /// <summary>
        ///Наименоваение объекта/договора
        /// </summary
        public string ContructionObjectCode { get; set; }
        /// <summary>
        ///Наименование подобъекта
        /// </summary
        public string ConstructionSubObjectCode { get; set; }
        /// <summary>
        /// Ссылка на родительскую модель
        /// </summary>
        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public MSGExellModel Owner { get; set; }
        /// <summary>
        /// Коллекия дочерних моделей
        /// </summary>
        public ObservableCollection<MSGExellModel> Children { get; set; } = new ObservableCollection<MSGExellModel>();

        private Excel.Worksheet _registerSheet;
        /// <summary>
        /// Прикрепленный к модели лист ведомости  Worksheet
        /// </summary>
        public Excel.Worksheet RegisterSheet
        {
            get
            {
                return _registerSheet;
            }
            set
            {
                if (!AllWorksheets.Contains(value))
                {
                    if (AllWorksheets.Contains(_registerSheet))
                        AllWorksheets.Remove(_registerSheet);
                    AllWorksheets.Add(value);
                }
                _registerSheet = value;

            }
        }

        private Excel.Worksheet _workerConsumptionsSheet;
        /// <summary>
        /// Прикрепленный к модели лист  Людских ресурсов Worksheet
        /// </summary>
        ///   
        public Excel.Worksheet WorkerConsumptionsSheet
        {
            get
            {
                return _workerConsumptionsSheet;
            }
            set
            {
                if (!AllWorksheets.Contains(value))
                {
                    if (AllWorksheets.Contains(_workerConsumptionsSheet))
                        AllWorksheets.Remove(_workerConsumptionsSheet);
                    AllWorksheets.Add(value);
                }
                _workerConsumptionsSheet = value;

            }
        }

        private Excel.Worksheet _machineConsumptionsSheet;
        /// <summary>
        /// Прикрепленный к модели лист  Технических ресурсов Worksheet
        /// </summary>
        ///   
        public Excel.Worksheet MachineConsumptionsSheet
        {
            get
            {
                return _machineConsumptionsSheet;
            }
            set
            {
                if (!AllWorksheets.Contains(value))
                {
                    if (AllWorksheets.Contains(_machineConsumptionsSheet))
                        AllWorksheets.Remove(_machineConsumptionsSheet);
                    AllWorksheets.Add(value);
                }
                _machineConsumptionsSheet = value;

            }
        }

        private Excel.Worksheet _commonSheet;

        /// <summary>
        /// Прикрепленный к модели лист общих данных  Worksheet
        /// </summary>
        public Excel.Worksheet CommonSheet
        {
            get
            {
                return _commonSheet;
            }
            set
            {
                if (!AllWorksheets.Contains(value))
                {
                    if (AllWorksheets.Contains(_commonSheet))
                        AllWorksheets.Remove(_commonSheet);
                    AllWorksheets.Add(value);
                }
                _commonSheet = value;
            }
        }
      
        /// <summary>
        /// Отвественных за работы отраженных в работах данной модели
        /// </summary>
        public Employer Employer { get; set; }
     
        public MSGExellModel()
        {
            WorksSections = new ExcelNotifyChangedCollection<WorksSection>();
            MSGWorks = new ExcelNotifyChangedCollection<MSGWork>();
            VOVRWorks = new ExcelNotifyChangedCollection<VOVRWork>();
            KSWorks = new ExcelNotifyChangedCollection<KSWork>();
            RCWorks = new ExcelNotifyChangedCollection<RCWork>();
            WorkReportCards = new ExcelNotifyChangedCollection<WorkReportCard>();
            WorkersComposition = new WorkersComposition();
            WorkerConsumptions = new WorkerConsumptions();
            MachinesComposition = new MachinesComposition();
            MachineConsumptions = new MachineConsumptions();
            UnitOfMeasurements = new ExcelNotifyChangedCollection<UnitOfMeasurement>();


        }
        /// <summary>
        /// Функция из части РАЗДЕЛЫ  листа Worksheet создает и помещает в модель  разделы работ
        /// </summary>
        public void LoadWorksSections()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            foreach (var section in this.WorksSections)
                this.Unregister(section);
            this.WorksSections.Clear();
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, WSEC_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    WorksSection w_section = new WorksSection();
                    this.Register(w_section, "Number", rowIndex, WSEC_NUMBER_COL, registerSheet);
                    this.Register(w_section, "Name", rowIndex, WSEC_NAME_COL, registerSheet);

                    w_section.Number = number.ToString();
                    if (this.WorksSections.FirstOrDefault(ws => ws.Number == w_section.Number) != null)
                        w_section.CellAddressesMap["Number"].IsValid = false;

                    var name = registerSheet.Cells[rowIndex, WSEC_NAME_COL].Value;
                    if (name != null)
                        w_section.Name = name;
                    else
                        w_section.CellAddressesMap["Name"].IsValid = false;

                    if (!this.WorksSections.Contains(w_section))
                        this.WorksSections.Add(w_section);
                }

                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из части МСГ листа Worksheet создает и помещает в модель работы типа MSGWork 
        /// </summary>
        public void LoadMSGWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;

            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            foreach (var work in this.MSGWorks)
                this.Unregister(work);
            this.MSGWorks.Clear();

            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, MSG_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    MSGWork msg_work = new MSGWork();

                    this.Register(msg_work, "Number", rowIndex, MSG_NUMBER_COL, this.RegisterSheet);
                    this.Register(msg_work, "Name", rowIndex, MSG_NAME_COL, this.RegisterSheet);
                    this.Register(msg_work, "ProjectQuantity", rowIndex, MSG_QUANTITY_COL, this.RegisterSheet);
                    this.Register(msg_work, "Quantity", rowIndex, MSG_QUANTITY_FACT_COL, this.RegisterSheet);
                    this.Register(msg_work, "Laboriousness", rowIndex, MSG_LABOURNESS_COL, this.RegisterSheet);
                    this.Register(msg_work, "UnitOfMeasurement.Name", rowIndex, MSG_MEASURE_COL, registerSheet);
                    var pr_quantity = registerSheet.Cells[rowIndex, MSG_QUANTITY_COL].Value;

                    msg_work.Number = number.ToString();
                    if (this.MSGWorks.FirstOrDefault(w => w.Number == msg_work.Number) != null)
                        msg_work.CellAddressesMap["Number"].IsValid = false;

                    var name = registerSheet.Cells[rowIndex, MSG_NAME_COL].Value;
                    if (name != null)
                        msg_work.Name = name;
                    else
                        msg_work.CellAddressesMap["Name"].IsValid = false;

                    var unit_of_measurement_name = registerSheet.Cells[rowIndex, MSG_MEASURE_COL].Value;
                    if (unit_of_measurement_name != null)
                        msg_work.UnitOfMeasurement = UnitOfMeasurements.FirstOrDefault(um => um.Name == unit_of_measurement_name.ToString());
                    else
                        msg_work.CellAddressesMap["UnitOfMeasurement.Name"].IsValid = false;

                    if (pr_quantity != null && pr_quantity != 0)
                        msg_work.ProjectQuantity = Decimal.Parse(pr_quantity.ToString());
                    else
                        msg_work.CellAddressesMap["ProjectQuantity"].IsValid = false;

                    var labourness = registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value;
                    if (labourness != null)
                        msg_work.Laboriousness = Decimal.Parse(labourness.ToString());
                    else
                        msg_work.CellAddressesMap["Laboriousness"].IsValid = false;

                    DateTime start_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                    DateTime end_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());

                    foreach (var sh in msg_work.WorkSchedules)
                        this.Unregister(sh);
                    msg_work.WorkSchedules.Clear();
                    WorkScheduleChunk work_sh_chunk = new WorkScheduleChunk(start_time, end_time);
                    int schedule_number = 1;
                    work_sh_chunk.Number = $"{msg_work.Number}.{schedule_number.ToString()}";
                    string is_snaday_vacation = registerSheet.Cells[rowIndex, MSG_SUNDAY_IS_VOCATION_COL].Value;
                    if (is_snaday_vacation != null && is_snaday_vacation.Contains("Нет"))
                        work_sh_chunk.IsSundayVacationDay = "Нет";
                    else
                        work_sh_chunk.IsSundayVacationDay = "Да";

                    this.Register(work_sh_chunk, "StartTime", rowIndex, MSG_START_DATE_COL, this.RegisterSheet);
                    this.Register(work_sh_chunk, "EndTime", rowIndex, MSG_END_DATE_COL, this.RegisterSheet);
                    this.Register(work_sh_chunk, "IsSundayVacationDay", rowIndex, MSG_SUNDAY_IS_VOCATION_COL, this.RegisterSheet);
                    msg_work.WorkSchedules.Add(work_sh_chunk);

                    while (registerSheet.Cells[rowIndex + 1, MSG_NUMBER_COL].Value == null
                                 && registerSheet.Cells[rowIndex + 1, MSG_START_DATE_COL].Value != null)
                    {
                        rowIndex++;
                        schedule_number++;
                        start_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                        end_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());
                        WorkScheduleChunk extra_work_sh_chunk = new WorkScheduleChunk(start_time, end_time);
                        extra_work_sh_chunk.Number = $"{msg_work.Number}.{schedule_number.ToString()}";

                        is_snaday_vacation = registerSheet.Cells[rowIndex, MSG_SUNDAY_IS_VOCATION_COL].Value;
                        if (is_snaday_vacation != null && is_snaday_vacation.Contains("Нет"))
                            extra_work_sh_chunk.IsSundayVacationDay = "Нет";
                        else
                            extra_work_sh_chunk.IsSundayVacationDay = "Да";
                        this.Register(extra_work_sh_chunk, "StartTime", rowIndex, MSG_START_DATE_COL, this.RegisterSheet);
                        this.Register(extra_work_sh_chunk, "EndTime", rowIndex, MSG_END_DATE_COL, this.RegisterSheet);
                        this.Register(extra_work_sh_chunk, "IsSundayVacationDay", rowIndex, MSG_SUNDAY_IS_VOCATION_COL, this.RegisterSheet);
                        msg_work.WorkSchedules.Add(extra_work_sh_chunk);
                    }

                    if (!this.MSGWorks.Contains(msg_work))
                        this.MSGWorks.Add(msg_work);
                    WorksSection w_section = this.WorksSections.Where(ws => ws.Number.StartsWith(msg_work.Number.Remove(msg_work.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (w_section != null)
                    {
                        MSGWork finded_msg_work = w_section.MSGWorks.FirstOrDefault(msgw => msgw.Number == msg_work.Number);
                        if (finded_msg_work == null)
                        {
                            w_section.MSGWorks.Add(msg_work);
                            msg_work.Owner = w_section;
                        }
                        else
                        {
                            finded_msg_work.CellAddressesMap["Number"].IsValid = false;
                            msg_work.CellAddressesMap["Number"].IsValid = false;
                        }
                    }
                }
                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из части МСГ листа Worksheet 
        /// </summary>
        public void LoadMSGWorkerCompositions()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;

            int rowIndex = FIRST_ROW_INDEX;
            foreach (var work_coposition in this.WorkersComposition)
                this.Unregister(work_coposition);
            this.WorkersComposition.Clear();
            foreach (var work in this.MSGWorks)
                work.WorkersComposition.Clear();

            null_str_count = 0;
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    NeedsOfWorker msg_needs_of_workers = new NeedsOfWorker();

                    this.Register(msg_needs_of_workers, "Number", rowIndex, MSG_NEEDS_OF_WORKERS_NUMBER_COL, this.RegisterSheet);
                    this.Register(msg_needs_of_workers, "Name", rowIndex, MSG_NEEDS_OF_WORKERS_NAME_COL, this.RegisterSheet);
                    this.Register(msg_needs_of_workers, "Quantity", rowIndex, MSG_NEEDS_OF_WORKERS_QUANTITY_COL, this.RegisterSheet);

                    msg_needs_of_workers.Number = number.ToString();
                    msg_needs_of_workers.Name = registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_NAME_COL].Value;

                    var quantity = registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_QUANTITY_COL].Value;
                    if (quantity != null)
                        msg_needs_of_workers.Quantity = Decimal.Parse(quantity.ToString());
                    else
                        msg_needs_of_workers.CellAddressesMap["Quantity"].IsValid = false;

                    MSGWork msg_work = this.MSGWorks.Where(w => w.Number.StartsWith(msg_needs_of_workers.Number.Remove(msg_needs_of_workers.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (msg_work != null)
                    {
                        msg_work.WorkersComposition.Add(msg_needs_of_workers);
                        msg_needs_of_workers.Owner = msg_work;
                        foreach (WorkScheduleChunk chunk in msg_work.WorkSchedules)
                        {
                            for (DateTime date = chunk.StartTime; date <= chunk.EndTime; date = date.AddDays(1))
                            {
                                NeedsOfWorkersDay needsOfWorkersDay = new NeedsOfWorkersDay();
                                needsOfWorkersDay.Date = date;
                                needsOfWorkersDay.Quantity = msg_needs_of_workers.Quantity;
                                msg_needs_of_workers.NeedsOfWorkersReportCard.Add(needsOfWorkersDay);
                            }
                        }
                    }

                    NeedsOfWorker global_needs_of_worker = this.WorkersComposition.FirstOrDefault(nw => nw.Name == msg_needs_of_workers.Name);
                    if (global_needs_of_worker == null)
                    {
                        global_needs_of_worker = new NeedsOfWorker();
                        global_needs_of_worker.Number = msg_needs_of_workers.Number;
                        global_needs_of_worker.Name = msg_needs_of_workers.Name;
                        foreach (NeedsOfWorkersDay needsOfWorkersDay in msg_needs_of_workers.NeedsOfWorkersReportCard)
                            global_needs_of_worker.NeedsOfWorkersReportCard.Add(needsOfWorkersDay);
                        this.WorkersComposition.Add(global_needs_of_worker);
                    }
                    else
                    {
                        foreach (NeedsOfWorkersDay needsOfWorkersDay in msg_needs_of_workers.NeedsOfWorkersReportCard)
                        {
                            var nw_day = global_needs_of_worker.NeedsOfWorkersReportCard.FirstOrDefault(nwd => nwd.Date == needsOfWorkersDay.Date);
                            if (nw_day != null)
                            {
                                nw_day.Quantity += needsOfWorkersDay.Quantity;
                            }
                            else
                            {
                                NeedsOfWorkersDay new_nw_day = new NeedsOfWorkersDay(needsOfWorkersDay);
                                global_needs_of_worker.NeedsOfWorkersReportCard.Add(new_nw_day);
                            }
                        }

                    }

                }
                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из части МСГ листа Worksheet 
        /// </summary>
        public void LoadMSGMachineCompositions()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;

            int rowIndex = FIRST_ROW_INDEX;
            foreach (var mch_coposition in this.MachinesComposition)
                this.Unregister(mch_coposition);
            this.MachinesComposition.Clear();
            foreach (var work in this.MSGWorks)
                work.MachinesComposition.Clear();

            null_str_count = 0;
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, MSG_NEEDS_OF_MACHINE_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    NeedsOfMachine msg_needs_of_machines = new NeedsOfMachine();

                    this.Register(msg_needs_of_machines, "Number", rowIndex, MSG_NEEDS_OF_MACHINE_NUMBER_COL, this.RegisterSheet);
                    this.Register(msg_needs_of_machines, "Name", rowIndex, MSG_NEEDS_OF_MACHINE_NAME_COL, this.RegisterSheet);
                    this.Register(msg_needs_of_machines, "Quantity", rowIndex, MSG_NEEDS_OF_MACHINE_QUANTITY_COL, this.RegisterSheet);

                    msg_needs_of_machines.Number = number.ToString();
                    msg_needs_of_machines.Name = registerSheet.Cells[rowIndex, MSG_NEEDS_OF_MACHINE_NAME_COL].Value;

                    var quantity = registerSheet.Cells[rowIndex, MSG_NEEDS_OF_MACHINE_QUANTITY_COL].Value;
                    if (quantity != null)
                        msg_needs_of_machines.Quantity = decimal.Parse(quantity.ToString());
                    else
                        msg_needs_of_machines.CellAddressesMap["Quantity"].IsValid = false;

                    MSGWork msg_work = this.MSGWorks.Where(w => w.Number.StartsWith(msg_needs_of_machines.Number.Remove(msg_needs_of_machines.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (msg_work != null)
                    {
                        msg_work.MachinesComposition.Add(msg_needs_of_machines);
                        msg_needs_of_machines.Owner = msg_work;
                        foreach (WorkScheduleChunk chunk in msg_work.WorkSchedules)
                        {
                            for (DateTime date = chunk.StartTime; date <= chunk.EndTime; date = date.AddDays(1))
                            {
                                NeedsOfMachineDay needsOfMachinesDay = new NeedsOfMachineDay();
                                needsOfMachinesDay.Date = date;
                                needsOfMachinesDay.Quantity = msg_needs_of_machines.Quantity;
                                msg_needs_of_machines.NeedsOfMachinesReportCard.Add(needsOfMachinesDay);
                            }
                        }
                    }

                    NeedsOfMachine global_needs_of_machine = this.MachinesComposition.FirstOrDefault(nw => nw.Name == msg_needs_of_machines.Name);
                    if (global_needs_of_machine == null)
                    {
                        global_needs_of_machine = new NeedsOfMachine();
                        global_needs_of_machine.Number = msg_needs_of_machines.Number;
                        global_needs_of_machine.Name = msg_needs_of_machines.Name;
                        foreach (NeedsOfMachineDay needsOfMachinesDay in msg_needs_of_machines.NeedsOfMachinesReportCard)
                            global_needs_of_machine.NeedsOfMachinesReportCard.Add(needsOfMachinesDay);
                        this.MachinesComposition.Add(global_needs_of_machine);
                    }
                    else
                    {
                        foreach (NeedsOfMachineDay needsOfMachinesDay in msg_needs_of_machines.NeedsOfMachinesReportCard)
                        {
                            var nw_day = global_needs_of_machine.NeedsOfMachinesReportCard.FirstOrDefault(nwd => nwd.Date == needsOfMachinesDay.Date);
                            if (nw_day != null)
                            {
                                nw_day.Quantity += needsOfMachinesDay.Quantity;
                            }
                            else
                            {
                                NeedsOfMachineDay new_nmch_day = new NeedsOfMachineDay(needsOfMachinesDay);
                                global_needs_of_machine.NeedsOfMachinesReportCard.Add(new_nmch_day);
                            }
                        }

                    }

                }
                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из части  ВОВР листа Worksheet создает и помещает в модель работы типа VOVRWork 
        /// </summary>
        public void LoadVOVRWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            foreach (var work in this.VOVRWorks)
                this.Unregister(work);
            this.VOVRWorks.Clear();
            foreach (var w in this.MSGWorks)
                w.VOVRWorks.Clear();

            null_str_count = 0;
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, VOVR_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    VOVRWork vovr_work = new VOVRWork();
                    this.Register(vovr_work, "Number", rowIndex, VOVR_NUMBER_COL, this.RegisterSheet);
                    this.Register(vovr_work, "Name", rowIndex, VOVR_NAME_COL, this.RegisterSheet);
                    this.Register(vovr_work, "ProjectQuantity", rowIndex, VOVR_QUANTITY_COL, this.RegisterSheet);
                    this.Register(vovr_work, "Quantity", rowIndex, VOVR_QUANTITY_FACT_COL, this.RegisterSheet);
                    this.Register(vovr_work, "Laboriousness", rowIndex, VOVR_LABOURNESS_COL, this.RegisterSheet);
                    this.Register(vovr_work, "UnitOfMeasurement.Name", rowIndex, VOVR_MEASURE_COL, this.RegisterSheet);

                    vovr_work.Number = number.ToString();
                    if (this.VOVRWorks.FirstOrDefault(w => w.Number == vovr_work.Number) != null)
                        vovr_work.CellAddressesMap["Number"].IsValid = false;

                    var name = registerSheet.Cells[rowIndex, VOVR_NAME_COL].Value;
                    if (name != null)
                        vovr_work.Name = name.ToString();
                    else
                        vovr_work.CellAddressesMap["Name"].IsValid = false;
                    var unit_of_measurement_name = registerSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value;
                    if (unit_of_measurement_name != null)
                        vovr_work.UnitOfMeasurement = new UnitOfMeasurement(unit_of_measurement_name);
                    else
                        vovr_work.CellAddressesMap["UnitOfMeasurement.Name"].IsValid = false;

                    var pr_quantity = registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL].Value;
                    if (pr_quantity != null && pr_quantity != 0)
                        vovr_work.ProjectQuantity = Decimal.Parse(pr_quantity.ToString());
                    else
                        vovr_work.CellAddressesMap["ProjectQuantity"].IsValid = false;

                    var labouriosness = registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL].Value;
                    if (labouriosness != null)
                        vovr_work.Laboriousness = Decimal.Parse(labouriosness.ToString());
                    else
                        vovr_work.CellAddressesMap["Laboriousness"].IsValid = false;

                    if (!this.VOVRWorks.Contains(vovr_work))
                        this.VOVRWorks.Add(vovr_work);

                    MSGWork msg_work = this.MSGWorks.Where(w => w.Number.StartsWith(vovr_work.Number.Remove(vovr_work.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (msg_work != null)
                    {
                        VOVRWork finded_vovr_wirk = msg_work.VOVRWorks.FirstOrDefault(vr_w => vr_w.Number == vovr_work.Number);
                        if (finded_vovr_wirk == null)
                        {
                            msg_work.VOVRWorks.Add(vovr_work);
                            vovr_work.Owner = msg_work;
                        }
                        else
                        {
                            finded_vovr_wirk.CellAddressesMap["Number"].IsValid = false;
                            vovr_work.CellAddressesMap["Number"].IsValid = false;
                        }
                    }

                }

                rowIndex++;
            }
        }

        /// <summary>
        /// Функция из части КС-2 листа Worksheet создает и помещает в модель работы типа KSWork 
        /// </summary>
        public void LoadKSWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            foreach (var work in this.KSWorks)
                this.Unregister(work);
            this.KSWorks.Clear();
            foreach (var w in this.VOVRWorks)
                w.KSWorks.Clear();
            null_str_count = 0;
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, KS_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    KSWork ks_work = new KSWork();
                    this.Register(ks_work, "Number", rowIndex, KS_NUMBER_COL, this.RegisterSheet);
                    this.Register(ks_work, "Code", rowIndex, KS_CODE_COL, this.RegisterSheet);
                    this.Register(ks_work, "Name", rowIndex, KS_NAME_COL, this.RegisterSheet);
                    this.Register(ks_work, "ProjectQuantity", rowIndex, KS_QUANTITY_COL, this.RegisterSheet);
                    this.Register(ks_work, "Quantity", rowIndex, KS_QUANTITY_FACT_COL, this.RegisterSheet);
                    this.Register(ks_work, "Laboriousness", rowIndex, KS_LABOURNESS_COL, this.RegisterSheet);
                    this.Register(ks_work, "UnitOfMeasurement.Name", rowIndex, KS_MEASURE_COL, this.RegisterSheet);

                    ks_work.Number = number;

                    var code = registerSheet.Cells[rowIndex, KS_CODE_COL].Value;
                    if (code != null)
                        ks_work.Code = code.ToString();
                    else
                        ks_work.CellAddressesMap["Code"].IsValid = false;

                    var name = registerSheet.Cells[rowIndex, KS_NAME_COL].Value;
                    if (name != null)
                        ks_work.Name = name.ToString();
                    else
                        ks_work.CellAddressesMap["Name"].IsValid = false;

                    var unit_of_measurement_name = registerSheet.Cells[rowIndex, KS_MEASURE_COL].Value;
                    if (unit_of_measurement_name != null)
                        ks_work.UnitOfMeasurement = new UnitOfMeasurement(unit_of_measurement_name);
                    else
                        ks_work.CellAddressesMap["UnitOfMeasurement.Name"].IsValid = false;

                    var pr_quantity = registerSheet.Cells[rowIndex, KS_QUANTITY_COL].Value;
                    if (pr_quantity != null && pr_quantity != 0)
                        ks_work.ProjectQuantity = Decimal.Parse(pr_quantity.ToString());
                    else
                        ks_work.CellAddressesMap["ProjectQuantity"].IsValid = false;

                    var laboriousness = registerSheet.Cells[rowIndex, KS_LABOURNESS_COL].Value;
                    if (laboriousness != null)
                        ks_work.Laboriousness = Decimal.Parse(laboriousness.ToString());
                    else
                        ks_work.CellAddressesMap["Laboriousness"].IsValid = false;


                    if (!this.KSWorks.Contains(ks_work))
                        this.KSWorks.Add(ks_work);

                    VOVRWork vovr_work = VOVRWorks.Where(w => w.Number.StartsWith(ks_work.Number.Remove(ks_work.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (vovr_work != null)
                    {
                        KSWork finded_ks_work = vovr_work.KSWorks.FirstOrDefault(kcw => kcw.Number == ks_work.Number);
                        if (finded_ks_work == null)
                        {
                            vovr_work.KSWorks.Add(ks_work);
                            ks_work.Owner = vovr_work;
                        }
                        else
                        {
                            finded_ks_work.CellAddressesMap["Number"].IsValid = false;
                            ks_work.CellAddressesMap["Number"].IsValid = false;
                        }


                    }
                }
                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из части КС-2 листа Worksheet создает и помещает в модель работы типа KSWork 
        /// </summary>
        public void LoadRCWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            foreach (var work in this.RCWorks)
                this.Unregister(work);
            this.RCWorks.Clear();
            foreach (var w in this.KSWorks)
                w.RCWorks.Clear();

            null_str_count = 0;
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, RC_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    RCWork rc_work = new RCWork();

                    this.Register(rc_work, "Number", rowIndex, RC_NUMBER_COL, this.RegisterSheet);
                    this.Register(rc_work, "Code", rowIndex, RC_CODE_COL, this.RegisterSheet);
                    this.Register(rc_work, "Name", rowIndex, RC_NAME_COL, this.RegisterSheet);
                    this.Register(rc_work, "ProjectQuantity", rowIndex, RC_QUANTITY_COL, this.RegisterSheet);
                    this.Register(rc_work, "Quantity", rowIndex, RC_QUANTITY_FACT_COL, this.RegisterSheet);
                    this.Register(rc_work, "LabournessCoefficient", rowIndex, RC_LABOURNESS_COEFFICIENT_COL, this.RegisterSheet);
                    this.Register(rc_work, "Laboriousness", rowIndex, RC_LABOURNESS_COL, this.RegisterSheet);
                    this.Register(rc_work, "UnitOfMeasurement.Name", rowIndex, RC_MEASURE_COL, this.RegisterSheet);


                    rc_work.Number = number;


                    var code = registerSheet.Cells[rowIndex, RC_CODE_COL].Value;
                    if (code != null)
                        rc_work.Code = code;
                    else
                        rc_work.CellAddressesMap["Code"].IsValid = false;

                    var name = registerSheet.Cells[rowIndex, RC_NAME_COL].Value;
                    if (name != null)
                        rc_work.Name = name;
                    else
                        rc_work.CellAddressesMap["Name"].IsValid = false;

                    var unit_of_measurement_name = registerSheet.Cells[rowIndex, RC_MEASURE_COL].Value;
                    if (unit_of_measurement_name != null)
                        rc_work.UnitOfMeasurement = new UnitOfMeasurement(unit_of_measurement_name);
                    else
                        rc_work.CellAddressesMap["UnitOfMeasurement.Name"].IsValid = false;

                    var pr_quantity = registerSheet.Cells[rowIndex, RC_QUANTITY_COL].Value;
                    if (pr_quantity != null && pr_quantity != 0)
                        rc_work.ProjectQuantity = Decimal.Parse(pr_quantity.ToString());
                    else
                        rc_work.CellAddressesMap["ProjectQuantity"].IsValid = false;

                    var laboriosness_coef = registerSheet.Cells[rowIndex, RC_LABOURNESS_COEFFICIENT_COL].Value;
                    if (laboriosness_coef != null)
                        rc_work.LabournessCoefficient = Decimal.Parse(laboriosness_coef.ToString());
                    else
                        rc_work.CellAddressesMap["LabournessCoefficient"].IsValid = false;

                    var laboriousness = registerSheet.Cells[rowIndex, RC_LABOURNESS_COL].Value;
                    if (laboriousness != null)
                        rc_work.Laboriousness = Decimal.Parse(laboriousness.ToString());
                    else
                        rc_work.CellAddressesMap["Laboriousness"].IsValid = false;

                    if (!this.RCWorks.Contains(rc_work))
                        this.RCWorks.Add(rc_work);
                    KSWork ks_work = this.KSWorks.Where(w => w.Number.StartsWith(rc_work.Number.Remove(rc_work.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (ks_work != null)
                    {
                        RCWork finded_rc_work = ks_work.RCWorks.FirstOrDefault(rcw => rcw.Number == rc_work.Number);
                        if (finded_rc_work == null)
                        {
                            ks_work.RCWorks.Add(rc_work);
                            rc_work.Owner = ks_work;
                        }
                        else
                        {
                            finded_rc_work.CellAddressesMap["Number"].IsValid = false;
                            rc_work.CellAddressesMap["Number"].IsValid = false;
                        }
                    }

                }
                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из календарной части (левой части) листа Worksheet создает и помещает в соответсвующие  
        /// работы типа KSWork табели выполненных работ ReportCard с объектами типа WorkDay с даной и количеством 
        /// выполенной работы
        /// </summary>
        public void LoadWorksReportCards()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            List<WorkReportCard> all_rcards = new List<WorkReportCard>(this.WorkReportCards);
            foreach (var rc in all_rcards)
            {
                this.WorkReportCards.Remove(rc);
                if (rc.Owner != null)
                    rc.Owner.ReportCard = null;
                this.Unregister(rc);
            }
            this.WorkReportCards.Clear();
            foreach (var w in this.RCWorks)
            {

                w.ReportCard = null;

            }
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, WRC_NUMBER_COL].Value;

                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    string rc_number = number;
                    WorkReportCard report_card = new WorkReportCard();

                    DateTime end_date = this.WorksEndDate; //DateTime.Parse(registerSheet.Cells[WORKS_END_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
                    report_card.Number = number;
                    this.Register(report_card, "Number", rowIndex, WRC_NUMBER_COL, this.RegisterSheet);
                    this.Register(report_card, "PreviousComplatedQuantity", rowIndex, WRC_PC_QUANTITY_COL, this.RegisterSheet);

                    var previus_comp_quantity = registerSheet.Cells[rowIndex, WRC_PC_QUANTITY_COL].Value;
                    if (previus_comp_quantity != null)
                        report_card.PreviousComplatedQuantity = Decimal.Parse(previus_comp_quantity.ToString());

                    int date_index = 0;
                    while (date_index < this.WorkedDaysNumber)
                    {
                        DateTime current_date = DateTime.Parse(registerSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString());
                        decimal quantity = 0;
                        if (registerSheet.Cells[rowIndex, WRC_DATE_COL + date_index].Value != null)
                            quantity = Decimal.Parse(registerSheet.Cells[rowIndex, WRC_DATE_COL + date_index].Value.ToString());
                        if (quantity != 0)
                        {
                            WorkDay workDay = new WorkDay();
                            workDay.Date = current_date;
                            workDay.Quantity = quantity;
                            this.Register(workDay, "Quantity", rowIndex, WRC_DATE_COL + date_index, this.RegisterSheet);
                            report_card.Add(workDay);
                        }
                        date_index++;
                    }
                    if (!this.WorkReportCards.Contains(report_card))
                        this.WorkReportCards.Add(report_card);
                    RCWork rc_work = this.RCWorks.FirstOrDefault(w => w.Number == rc_number);
                    if (rc_work != null)
                    {
                        rc_work.ReportCard = report_card;
                        report_card.Owner = rc_work;
                    }
                    else
                    {
                        if (rc_work != null && rc_work.ReportCard != null) rc_work.ReportCard.CellAddressesMap["Number"].IsValid = false;
                        report_card.CellAddressesMap["Number"].IsValid = false;
                    }

                }
                rowIndex++;
            }

        }
        public void LoadWorkerConsumptions()
        {
            Excel.Worksheet consumtionsSheet = this.WorkerConsumptionsSheet;
            int rowIndex = W_CONSUMPTIONS_FIRST_ROW_INDEX;
            foreach (var wc in this.WorkerConsumptions)
                this.Unregister(wc);
            this.WorkerConsumptions.Clear();
            null_str_count = 0;

            while (null_str_count < 100)
            {
                var number = consumtionsSheet.Cells[rowIndex, W_CONSUMPTIONS_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    WorkerConsumption worker_consumption = new WorkerConsumption();
                    this.Register(worker_consumption, "Number", rowIndex, W_CONSUMPTIONS_NUMBER_COL, consumtionsSheet);
                    this.Register(worker_consumption, "Name", rowIndex, W_CONSUMPTIONS_NAME_COL, consumtionsSheet);

                    worker_consumption.Number = number.ToString();
                    var name = consumtionsSheet.Cells[rowIndex, W_CONSUMPTIONS_NAME_COL].Value;
                    worker_consumption.Name = name;
                    worker_consumption.WorkersConsumptionReportCard.Clear();

                    int date_index = 0;

                    while (date_index < this.WorkedDaysNumber)
                    {
                        DateTime current_date = DateTime.Parse(consumtionsSheet.Cells[W_CONSUMPTIONS_DATE_RAW, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString());
                        decimal quantity = 0;
                        if (consumtionsSheet.Cells[rowIndex, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value != null)
                            quantity = Decimal.Parse(consumtionsSheet.Cells[rowIndex, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString());

                        if (quantity != 0)
                        {
                            WorkerConsumptionDay w_consumption_Day = new WorkerConsumptionDay();
                            w_consumption_Day.Date = current_date;
                            w_consumption_Day.Quantity = quantity;
                            this.Register(w_consumption_Day, "Quantity", rowIndex, W_CONSUMPTIONS_FIRST_DATE_COL + date_index, consumtionsSheet);
                            worker_consumption.WorkersConsumptionReportCard.Add(w_consumption_Day);
                        }

                        date_index++;
                    }
                    if (!this.WorkerConsumptions.Contains(worker_consumption))
                        this.WorkerConsumptions.Add(worker_consumption);

                }
                rowIndex++;
            }

        }
        public void LoadMachineConsumptions()
        {
            Excel.Worksheet consumtionsSheet = this.MachineConsumptionsSheet;
            int rowIndex = MCH_CONSUMPTIONS_FIRST_ROW_INDEX;
            foreach (var mc in this.MachineConsumptions)
                this.Unregister(mc);
            this.MachineConsumptions.Clear();
            null_str_count = 0;

            while (null_str_count < 100)
            {
                var number = consumtionsSheet.Cells[rowIndex, MCH_CONSUMPTIONS_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    MachineConsumption machine_consumption = new MachineConsumption();
                    this.Register(machine_consumption, "Number", rowIndex, MCH_CONSUMPTIONS_NUMBER_COL, consumtionsSheet);
                    this.Register(machine_consumption, "Name", rowIndex, MCH_CONSUMPTIONS_NAME_COL, consumtionsSheet);

                    machine_consumption.Number = number.ToString();
                    var name = consumtionsSheet.Cells[rowIndex, MCH_CONSUMPTIONS_NAME_COL].Value;
                    machine_consumption.Name = name;
                    machine_consumption.MachinesConsumptionReportCard.Clear();

                    int date_index = 0;

                    while (date_index < this.WorkedDaysNumber)
                    {
                        DateTime current_date = DateTime.Parse(consumtionsSheet.Cells[MCH_CONSUMPTIONS_DATE_RAW, MCH_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString());
                        decimal quantity = 0;
                        if (consumtionsSheet.Cells[rowIndex, MCH_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value != null)
                            quantity = Decimal.Parse(consumtionsSheet.Cells[rowIndex, MCH_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString());

                        if (quantity != 0)
                        {
                            MachineConsumptionDay w_consumption_Day = new MachineConsumptionDay();
                            w_consumption_Day.Date = current_date;
                            w_consumption_Day.Quantity = quantity;
                            this.Register(w_consumption_Day, "Quantity", rowIndex, MCH_CONSUMPTIONS_FIRST_DATE_COL + date_index, consumtionsSheet);
                            machine_consumption.MachinesConsumptionReportCard.Add(w_consumption_Day);
                        }

                        date_index++;
                    }
                    if (!this.MachineConsumptions.Contains(machine_consumption))
                        this.MachineConsumptions.Add(machine_consumption);

                }
                rowIndex++;
            }

        }
        /// <summary>
        /// Заргужает(перезагружает)  данныхе из соотвествующих листов в модель
        /// </summary>
        public void ReloadSheetModel()
        {
            this.UpdateCellAddressMapsWorkSheets();
            this.WorksStartDate = DateTime.Parse(this.RegisterSheet.Cells[WORKS_START_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
            this.ContractCode = this.CommonSheet.Cells[CONTRACT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
            this.ContructionObjectCode = this.CommonSheet.Cells[CONSTRUCTION_OBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
            this.ConstructionSubObjectCode = this.CommonSheet.Cells[CONSTRUCTION_SUBOBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
            //this.CellAddressesMap.Add("ContractCode", new ExellPropAddress<int, int, Worksheet>(CONTRACT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));
            //this.CellAddressesMap.Add("ContructionObjectCode", new ExellPropAddress<int, int, Worksheet>(CONSTRUCTION_OBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));
            //this.CellAddressesMap.Add("ConstructionSubObjectCode", new ExellPropAddress<int, int, Worksheet>(CONSTRUCTION_SUBOBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));

            this.WorksStartDate = DateTime.Parse(this.RegisterSheet.Cells[WORKS_START_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());

            if (this.Owner == null)
            {
                this.LoadWorksSections();
                this.LoadMSGWorks();
                this.LoadMSGWorkerCompositions();
                this.LoadMSGMachineCompositions();
                this.LoadVOVRWorks();
                this.LoadKSWorks();
                this.LoadRCWorks();
                this.LoadWorksReportCards();
                this.LoadWorkerConsumptions();
                this.LoadMachineConsumptions();
                foreach (MSGExellModel model in Children)
                    model.ReloadSheetModel();

            }
            else
            {
                this.CopyOwnerObjectModels();
                this.LoadWorksReportCards();
                this.LoadWorkerConsumptions();
                this.LoadMachineConsumptions();
            }
        }
        /// <summary>
        /// Функция форматирует представления модели на листе Excel
        /// </summary>
        public void SetStyleFormats()
        {
            //  this.UpdateCellAddressMapsWorkSheets();
            this.RemoveGroups(this.RegisterSheet);
            int selectin_col = W_SECTION_COLOR;
            this.SetBordersBoldLine(this.WorksSections.GetRange(this.RegisterSheet), XlLineStyle.xlLineStyleNone, XlLineStyle.xlDashDot, XlLineStyle.xlLineStyleNone, XlLineStyle.xlLineStyleNone);
            foreach (WorksSection section in this.WorksSections)
                this.SetStyleFormats( section,  selectin_col);

            this.WorksSections.SetInvalidateCellsColor(XlRgbColor.rgbRed);
            this.MSGWorks.SetInvalidateCellsColor(XlRgbColor.rgbRed);
            this.VOVRWorks.SetInvalidateCellsColor(XlRgbColor.rgbRed);
            this.KSWorks.SetInvalidateCellsColor(XlRgbColor.rgbRed);
            this.RCWorks.SetInvalidateCellsColor(XlRgbColor.rgbRed);
            this.WorkReportCards.SetInvalidateCellsColor(XlRgbColor.rgbRed);

            this.SetBordersBoldLine(this.WorkerConsumptions.GetRange(this.WorkerConsumptionsSheet));
            int w_consumption_col = 33;
            foreach (WorkerConsumption consumption in this.WorkerConsumptions)
            {
                // consumption.GetRange(this.WorkerConsumptionsSheet).Interior.ColorIndex = w_consumption_col++;
                int days_namber = (this.WorksEndDate - this.WorksStartDate).Days;
                Excel.Range cons_range = this.WorkerConsumptionsSheet.Range[
                    this.WorkerConsumptionsSheet.Cells[consumption.CellAddressesMap["Number"].Row, W_CONSUMPTIONS_NUMBER_COL],
                    this.WorkerConsumptionsSheet.Cells[consumption.CellAddressesMap["Number"].Row, days_namber]];
                cons_range.Interior.ColorIndex = w_consumption_col++;
                cons_range.Borders.LineStyle = XlLineStyle.xlDashDotDot;
                this.SetBordersBoldLine(cons_range, XlLineStyle.xlDouble, XlLineStyle.xlDouble,
                    XlLineStyle.xlContinuous, XlLineStyle.xlContinuous);
            }

            Excel.Range vovr_colomns = this.RegisterSheet.Range[this.RegisterSheet.Columns[VOVR_NUMBER_COL], this.RegisterSheet.Columns[VOVR_LABOURNESS_COL]];
            Excel.Range ks_colomns = this.RegisterSheet.Range[this.RegisterSheet.Columns[VOVR_NUMBER_COL], this.RegisterSheet.Columns[KS_LABOURNESS_COL]];
            try
            {
                ks_colomns.Group();
                vovr_colomns.Group();
            }
            catch
            {

            }

        }
        public int  SetStyleFormats(WorksSection section, int selectin_col)
        {

            var section_range = section.GetRange(this.RegisterSheet);
            section_range.Interior.ColorIndex = selectin_col;
            this.SetBordersBoldLine(section_range);
           int   first_row = section_range.Row;
            this.SetBordersBoldLine(section.MSGWorks.GetRange(this.RegisterSheet), XlLineStyle.xlLineStyleNone, XlLineStyle.xlDashDot, XlLineStyle.xlLineStyleNone, XlLineStyle.xlLineStyleNone);
            int msg_work_col = selectin_col + 1;
            int last_section_row = 0;
            foreach (MSGWork msg_work in section.MSGWorks)
            {
                last_section_row = this.SetStyleFormats(msg_work, msg_work_col);
            }
            try
            {
                Excel.Range range = this.RegisterSheet.Range[this.RegisterSheet.Rows[first_row + 1],
                                    this.RegisterSheet.Rows[last_section_row + _SECTIONS_GAP]];
                range.Group();
            }
            catch
            {

            }
            return last_section_row;
        }

        public int  SetStyleFormats(MSGWork msg_work, int msg_work_col)
        {
            var msg_work_range = msg_work.GetRange(this.RegisterSheet, MSG_LABOURNESS_COL);
            msg_work_range.Interior.ColorIndex = msg_work_col;
            this.SetBordersBoldLine(msg_work_range);
            int last_msg_row = msg_work_range.Row;

            this.SetBordersBoldLine(msg_work.WorkersComposition.GetRange(this.RegisterSheet));
            int need_of_workers_count = 0;
            foreach (NeedsOfWorker need_of_worker in msg_work.WorkersComposition)
            {
                var need_of_worker_range = need_of_worker.GetRange(this.RegisterSheet);
                need_of_worker_range.Interior.ColorIndex = msg_work_col;
                need_of_workers_count++;
            }

            this.SetBordersBoldLine(msg_work.MachinesComposition.GetRange(this.RegisterSheet));

            int need_of_machine_count = 0;
            foreach (NeedsOfMachine need_of_machine in msg_work.MachinesComposition)
            {
                var need_of_machine_range = need_of_machine.GetRange(this.RegisterSheet);
                need_of_machine_range.Interior.ColorIndex = msg_work_col;
                need_of_machine_count++;
            }

            this.SetBordersBoldLine(msg_work.WorkSchedules.GetRange(this.RegisterSheet));
            int chunks_count = 0;
            foreach (WorkScheduleChunk chunk in msg_work.WorkSchedules)
            {
                var work_composition_range = chunk.GetRange(this.RegisterSheet);
                work_composition_range.Interior.ColorIndex = msg_work_col;
                chunks_count++;
            }

            this.SetBordersBoldLine(msg_work.VOVRWorks.GetRange(this.RegisterSheet), XlLineStyle.xlLineStyleNone, XlLineStyle.xlDashDot, XlLineStyle.xlLineStyleNone, XlLineStyle.xlLineStyleNone);
            int vovr_work_col = msg_work_col + 1;
            foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
            {
                int  vovr_work_row = this.SetStyleFormats(vovr_work, vovr_work_col);
                last_msg_row++;
                vovr_work_col++;
            }

            if (last_msg_row < need_of_workers_count) last_msg_row = need_of_workers_count;
            if (last_msg_row < need_of_machine_count) last_msg_row = need_of_machine_count;
            if (last_msg_row < chunks_count) last_msg_row = chunks_count;
            return last_msg_row;

        }


        public  int  SetStyleFormats(VOVRWork vovr_work, int vovr_work_col)
        {
            var vovr_work_range = vovr_work.GetRange(this.RegisterSheet);
            vovr_work_range.Interior.ColorIndex = vovr_work_col;
            this.SetBordersBoldLine(vovr_work_range);
            this.SetBordersBoldLine(vovr_work.KSWorks.GetRange(this.RegisterSheet), XlLineStyle.xlLineStyleNone, XlLineStyle.xlDashDot, XlLineStyle.xlLineStyleNone, XlLineStyle.xlLineStyleNone);
            int ks_work_col = vovr_work_col;
            int last_vovr_row = vovr_work_range.Row;
            foreach (KSWork ks_work in vovr_work.KSWorks)
            {
                int  ks_work_row = SetStyleFormats(ks_work, ks_work_col);
                if (last_vovr_row < ks_work_row) last_vovr_row = ks_work_row;
            }
            try
            {
                Excel.Range top_row = this.RegisterSheet.Rows[vovr_work.KSWorks.GetTopRow() + 1];
                Excel.Range rottom_row_num = this.RegisterSheet.Rows[vovr_work.KSWorks.OrderBy(w => w.RCWorks.GetBottomRow()).Last().RCWorks.GetBottomRow()]; ;
                this.RegisterSheet.Range[top_row, rottom_row_num].Group();
            }
            catch { }
            return last_vovr_row;
        }


        public int  SetStyleFormats(KSWork ks_work, int ks_work_col)
        {
            var ks_work_range = ks_work.GetRange(this.RegisterSheet);
            ks_work_range.Interior.ColorIndex = ks_work_col;
            int last_row = ks_work_range.Row;
            this.SetBordersBoldLine(ks_work.RCWorks.GetRange(this.RegisterSheet, RC_LABOURNESS_COL));
            foreach (RCWork rc_work in ks_work.RCWorks)
            {
                int rc_row = this.SetStyleFormats(rc_work, ks_work_col);
                if (last_row < rc_row) last_row = rc_row;
            }

            if (ks_work.RCWorks.Count > 0)
            {
                int rc_works_top_row = ks_work.RCWorks.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).First().Value.Row;
                int rc_works_bottom_row = ks_work.RCWorks.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).Last().Value.Row;
                int days_number = (this.WorksEndDate - this.WorksStartDate).Days;
                var report_cards_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[rc_works_top_row, WRC_NUMBER_COL],
                                                                this.RegisterSheet.Cells[rc_works_bottom_row, WRC_PC_QUANTITY_COL + days_number]];
                this.SetBordersBoldLine(report_cards_range, XlLineStyle.xlContinuous);

            }
            this.SetBordersBoldLine(ks_work_range, XlLineStyle.xlLineStyleNone);

            return last_row;
        }
        public  int  SetStyleFormats(RCWork rc_work, int ks_work_col)
        {
            var rc_work_range = rc_work.GetRange(this.RegisterSheet, RC_LABOURNESS_COL);
            rc_work_range.Interior.ColorIndex = ks_work_col;
            if (rc_work.ReportCard != null && rc_work.ReportCard.CellAddressesMap.Count > 0)
            {
                var cr_range = rc_work.ReportCard.GetRange(this.RegisterSheet);
                if (cr_range != null)
                {
                    cr_range.Interior.ColorIndex = ks_work_col;
                    // cr_range.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    this.SetBordersBoldLine(cr_range, XlLineStyle.xlDashDotDot, XlLineStyle.xlDashDotDot, XlLineStyle.xlContinuous, XlLineStyle.xlContinuous);
                }
                // Excel.Range last_cell = this.RegisterSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

                var days_row_range = this.RegisterSheet.Range[
                       this.RegisterSheet.Cells[rc_work.ReportCard.CellAddressesMap["Number"].Cell.Row, WRC_PC_QUANTITY_COL + 1],
                       this.RegisterSheet.Cells[rc_work.ReportCard.CellAddressesMap["Number"].Cell.Row, WRC_PC_QUANTITY_COL + this.WorkedDaysNumber]];
                days_row_range.Interior.ColorIndex = ks_work_col;
                days_row_range.Borders.LineStyle = Excel.XlLineStyle.xlDashDotDot;

                this.SetBordersBoldLine(days_row_range,
                    XlLineStyle.xlContinuous, XlLineStyle.xlContinuous,
                    XlLineStyle.xlContinuous, XlLineStyle.xlContinuous);

            }
            return rc_work_range.Row;
        }

        /// <summary>
        ///Фунция проставляет все соотвесвующие формулы в ячейках Excell в соотвествии с моделью
        /// </summary>
        public void SetFormulas()
        {
            int days_number = (this.WorksEndDate - this.WorksStartDate).Days;

            // Excel.Range tmp_first_rc_work_quantity_cell = null;
            Excel.Range tmp_first_rc_card_days_row = null;
            if (this.Owner == null && this.WorksSections.Count > 0
                && this.WorksSections[0].MSGWorks.Count > 0
                && this.WorksSections[0].MSGWorks[0].VOVRWorks.Count > 0
                && this.WorksSections[0].MSGWorks[0].VOVRWorks[0].KSWorks.Count > 0
                && this.WorksSections[0].MSGWorks[0].VOVRWorks[0].KSWorks[0].RCWorks.Count > 0)
            {
                RCWork first_rc_work = this.WorksSections[0].MSGWorks[0].VOVRWorks[0].KSWorks[0].RCWorks[0];
                Excel.Range first_cell = this.RegisterSheet.Cells[first_rc_work.CellAddressesMap["Number"].Row, WRC_PC_QUANTITY_COL];
                Excel.Range last_cell = this.RegisterSheet.Cells[first_rc_work.CellAddressesMap["Number"].Row, WRC_PC_QUANTITY_COL + days_number + 1];

                string first_rc_record_day_formula = "";
                Excel.Range tmp_first_rc_record_day_range = first_cell;
                foreach (MSGExellModel model in this.Children)
                {
                    first_rc_record_day_formula += $"{model.RegisterSheet.Name}!{Func.RangeAddress(model.RegisterSheet.Cells[first_cell.Row, first_cell.Column])}+";
                }
                first_rc_record_day_formula = first_rc_record_day_formula.TrimEnd('+');
                if (first_rc_record_day_formula != "")
                    tmp_first_rc_record_day_range.Formula = $"={first_rc_record_day_formula}";

                int date_iterator = 0;
                tmp_first_rc_record_day_range.Copy();
                while (date_iterator <= days_number + 1)
                {
                    this.RegisterSheet.Cells[first_cell.Row, first_cell.Column + date_iterator].PasteSpecial(XlPasteType.xlPasteAll);
                    date_iterator++;
                }
                tmp_first_rc_card_days_row = this.RegisterSheet.Range[first_cell, last_cell];

            }

            foreach (WorksSection section in this.WorksSections)
            {
                foreach (MSGWork msg_work in section.MSGWorks)
                {
                    //foreach (NeedsOfWorker need_of_worker in msg_work.WorkersComposition)
                    //{
                    //}
                    string msg_works_labourness_sum_formula = "";
                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        string vovr_works_labourness_sum_formula = "";
                        foreach (KSWork ks_work in vovr_work.KSWorks)
                        {
                            //string rc_works_labourness_sum_formula = "";
                            //var first_cell = this.RegisterSheet.Cells[section.CellAddressesMap["Number"].Row, WRC_PC_QUANTITY_COL];
                            //var lastt_cell = this.RegisterSheet.Cells[section.CellAddressesMap["Number"].Row, WRC_PC_QUANTITY_COL + 1 + (this.WorksEndDate - this.WorksStartDate).Days];

                            //Excel.Range q_summ_range = this.RegisterSheet.Cells[section.CellAddressesMap["Number"].Row, RC_QUANTITY_COL];
                            //q_summ_range.Formula = $"=SUM({Func.RangeAddress(first_cell)}:{Func.RangeAddress(lastt_cell)})";

                            string rc_works_labourness_sum_formula = "";
                            if (this.Owner == null) tmp_first_rc_card_days_row.Copy();

                            foreach (RCWork rc_work in ks_work.RCWorks)
                            {
                                if (rc_work.ReportCard == null)
                                {
                                    rc_work.ReportCard = new WorkReportCard();
                                    this.Register(rc_work.ReportCard, "Number", rc_work.CellAddressesMap["Number"].Row, WRC_NUMBER_COL, this.RegisterSheet);
                                    this.Register(rc_work.ReportCard, "PreviousComplatedQuantity", rc_work.CellAddressesMap["Number"].Row, WRC_PC_QUANTITY_COL, this.RegisterSheet);
                                    rc_work.ReportCard.Number = rc_work.Number;
                                }

                                var first_cell = this.RegisterSheet.Cells[rc_work.CellAddressesMap["Number"].Row, WRC_PC_QUANTITY_COL];
                                var lastt_cell = this.RegisterSheet.Cells[rc_work.CellAddressesMap["Number"].Row, WRC_PC_QUANTITY_COL + 1 + days_number];
                                Excel.Range q_summ_range = this.RegisterSheet.Cells[rc_work.CellAddressesMap["Number"].Row, RC_QUANTITY_FACT_COL];
                                q_summ_range.Formula = $"=SUM({Func.RangeAddress(first_cell)}:{Func.RangeAddress(lastt_cell)})";

                                if (this.Owner == null)
                                {
                                    Excel.Range w_days_row_range = this.RegisterSheet.Cells[rc_work.ReportCard.CellAddressesMap["Number"].Row, WRC_PC_QUANTITY_COL];
                                    if (tmp_first_rc_card_days_row != null)
                                        w_days_row_range.PasteSpecial(XlPasteType.xlPasteAll);
                                }
                                rc_works_labourness_sum_formula +=
                                       $"{Func.RangeAddress(rc_work.CellAddressesMap["Quantity"].Cell)}*{Func.RangeAddress(rc_work.CellAddressesMap["Laboriousness"].Cell)}+";


                            }
                            rc_works_labourness_sum_formula = rc_works_labourness_sum_formula.TrimEnd('+');
                            if (rc_works_labourness_sum_formula != "")
                            {
                                string ks_quantity_formula = $"=({rc_works_labourness_sum_formula})/{Func.RangeAddress(ks_work.CellAddressesMap["Laboriousness"].Cell)}";
                                ks_work.CellAddressesMap["Quantity"].Cell.Formula = ks_quantity_formula;

                                vovr_works_labourness_sum_formula +=
                                    $"{Func.RangeAddress(ks_work.CellAddressesMap["Quantity"].Cell)}*{Func.RangeAddress(ks_work.CellAddressesMap["Laboriousness"].Cell)}+";
                            }
                        }
                        vovr_works_labourness_sum_formula = vovr_works_labourness_sum_formula.TrimEnd('+');

                        if (vovr_works_labourness_sum_formula != "")
                        {
                            string vovr_quantity_formula = $"=({vovr_works_labourness_sum_formula})/{Func.RangeAddress(vovr_work.CellAddressesMap["Laboriousness"].Cell)}";
                            vovr_work.CellAddressesMap["Quantity"].Cell.Formula = vovr_quantity_formula;
                        }
                        msg_works_labourness_sum_formula +=
                                             $"{Func.RangeAddress(vovr_work.CellAddressesMap["Quantity"].Cell)}*{Func.RangeAddress(vovr_work.CellAddressesMap["Laboriousness"].Cell)}+";

                    }
                    msg_works_labourness_sum_formula = msg_works_labourness_sum_formula.TrimEnd('+');
                    if (msg_works_labourness_sum_formula != "")
                    {
                        string msg_quantity_formula = $"=({msg_works_labourness_sum_formula})/{Func.RangeAddress(msg_work.CellAddressesMap["Laboriousness"].Cell)}";
                        msg_work.CellAddressesMap["Quantity"].Cell.Formula = msg_quantity_formula;
                    }

                }

            }

            foreach (WorkerConsumption consumption in this.WorkerConsumptions)
            {
                int col_iterator = W_CONSUMPTIONS_FIRST_DATE_COL;
                while (col_iterator <= (this.WorksEndDate - this.WorksStartDate).Days)
                {
                    var cons_day_range = this.WorkerConsumptionsSheet.Cells[consumption.CellAddressesMap["Number"].Row, col_iterator];
                    string cons_quantity_formula = "";
                    foreach (MSGExellModel model in this.Children)
                    {
                        var child_consumption = model.WorkerConsumptions.FirstOrDefault(cn => cn.Number == consumption.Number);
                        if (child_consumption != null)
                        {
                            int cons_row = child_consumption.CellAddressesMap["Number"].Row;
                            var child_cons_day_range =
                                 model.WorkerConsumptionsSheet.Cells[cons_row, col_iterator];
                            cons_quantity_formula += $"{model.WorkerConsumptionsSheet.Name}!{Func.RangeAddress(cons_day_range)}+";
                        }
                    }
                    cons_quantity_formula = cons_quantity_formula.TrimEnd('+');
                    if (cons_quantity_formula != "")
                        cons_day_range.Formula = $"={cons_quantity_formula}";
                    col_iterator++;
                }

            }
            foreach (MachineConsumption consumption in this.MachineConsumptions)
            {
                int col_iterator = MCH_CONSUMPTIONS_FIRST_DATE_COL;
                while (col_iterator <= (this.WorksEndDate - this.WorksStartDate).Days)
                {
                    var cons_day_range = this.MachineConsumptionsSheet.Cells[consumption.CellAddressesMap["Number"].Row, col_iterator];
                    string cons_quantity_formula = "";
                    foreach (MSGExellModel model in this.Children)
                    {
                        var child_consumption = model.MachineConsumptions.FirstOrDefault(cn => cn.Number == consumption.Number);
                        if (child_consumption != null)
                        {
                            int cons_row = child_consumption.CellAddressesMap["Number"].Row;
                            var child_cons_day_range =
                                 model.MachineConsumptionsSheet.Cells[cons_row, col_iterator];
                            cons_quantity_formula += $"{model.MachineConsumptionsSheet.Name}!{Func.RangeAddress(cons_day_range)}+";
                        }
                    }
                    cons_quantity_formula = cons_quantity_formula.TrimEnd('+');
                    if (cons_quantity_formula != "")
                        cons_day_range.Formula = $"={cons_quantity_formula}";
                    col_iterator++;
                }

            }
        }
        /// <summary>
        /// Функция устанавливает границы диапазона двоейной линией
        /// </summary>
        /// <param name="range"></param>
        private void SetBordersBoldLine(Excel.Range range)
        {
            if (range == null) return;
            //range.Borders.LineStyle = Excel.XlLineStyle.xlDot;
            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlDouble;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlDouble;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDouble;
        }
        /// <summary>
        /// Функция устанавливает границы диапазона двойной линей
        /// </summary>
        /// <param name="range"></param>
        /// <param name="right"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="bottom"></param>
        private void SetBordersBoldLine(Excel.Range range, bool right = true, bool left = true, bool top = true, bool bottom = true)
        {
            if (range == null) return;

            if (left) range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            if (top) range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            if (bottom) range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            if (right) range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }
        /// <summary>
        /// Функция устанавливает границы диапазона соовествующими типами линий
        /// </summary>
        /// <param name="range"></param>
        /// <param name="right"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="bottom"></param>
        private void SetBordersBoldLine(Excel.Range range,
            Excel.XlLineStyle right = Excel.XlLineStyle.xlDouble,
            Excel.XlLineStyle left = Excel.XlLineStyle.xlDouble,
            Excel.XlLineStyle top = Excel.XlLineStyle.xlDouble,
            Excel.XlLineStyle bottom = Excel.XlLineStyle.xlDouble)
        {
            if (range == null) return;

            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = left;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = top;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = bottom;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = right;
        }
        //private void ClearStyleFormatsByObjects()
        //{
        //    this.GetRange(this.RegisterSheet).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        //    this.GetRange(this.RegisterSheet).Interior.ColorIndex = 0;

        //}
        /// <summary>
        /// Функиця пересчета трудоемкостей всех типов работ исходя из проставленных в трудоемкостей
        /// в работах типа КС-2
        /// </summary>
        public void CalcLabourness()
        {
            foreach (WorksSection section in this.WorksSections)
            {
                foreach (MSGWork msg_work in section.MSGWorks)
                {
                    //  if (msg_work.Laboriousness == 0)
                    {
                        decimal common_vovr_laboueness = 0;
                        foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                        {
                            // if (vovr_work.Laboriousness == 0)
                            {
                                decimal common_ks_laboueness = 0;
                                foreach (KSWork ks_work in vovr_work.KSWorks)
                                {
                                    common_ks_laboueness += ks_work.ProjectQuantity * ks_work.Laboriousness;

                                    this.CalcLabournessCoefficiens(ks_work);
                                    foreach (RCWork rc_work in ks_work.RCWorks)
                                        rc_work.Laboriousness = ks_work.ProjectQuantity * ks_work.Laboriousness * rc_work.LabournessCoefficient / rc_work.ProjectQuantity;
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
        /// <summary>
        /// Функция вычисляет коэфиценты рапределения трудоемкойстей для работ. 
        /// </summary>
        /// <param name="ks_work"></param>
        /// <exception cref="Exception"></exception>
        private void CalcLabournessCoefficiens(KSWork ks_work)
        {
            var rc_works_with_notNull_labourness = ks_work.RCWorks.Where(rcw => rcw.Laboriousness != 0);
            decimal rc_laboriousness_coeffecients_sum = 0;
            var ks_work_total_laboriousnes = (ks_work.Laboriousness * ks_work.ProjectQuantity);
            foreach (RCWork rc_work in rc_works_with_notNull_labourness)
                rc_work.LabournessCoefficient = rc_work.Laboriousness * rc_work.ProjectQuantity / ks_work_total_laboriousnes;

            var rc_works_with_notNull_labourness_coef = ks_work.RCWorks.Where(rcw => rcw.LabournessCoefficient != 0);
            foreach (RCWork rc_work in rc_works_with_notNull_labourness_coef)
                rc_laboriousness_coeffecients_sum += rc_work.LabournessCoefficient;

            var rc_works_with_Null_labourness_coef = ks_work.RCWorks.Where(rcw => rcw.LabournessCoefficient == 0).ToList();
            foreach (RCWork rc_work in rc_works_with_Null_labourness_coef)
            {
                decimal coef = (1 - rc_laboriousness_coeffecients_sum) / rc_works_with_Null_labourness_coef.ToList().Count;
                if (coef <= 0)
                {

                    rc_work.CellAddressesMap[nameof(rc_work.LabournessCoefficient)].IsValid = false;
                    throw new Exception("Кофицент должен быть больше нуля!");
                }
                rc_work.LabournessCoefficient = coef;
            }


        }
        /// <summary>
        /// Функцич подсчета объемов выполненных работ 
        /// </summary>
        public void CalcQuantity()
        {
            List<MSGExellModel> loaded_models = new List<MSGExellModel>();
            foreach (WorksSection section in this.WorksSections)
            {
                foreach (MSGWork msg_work in section.MSGWorks)
                {

                    var msg_work_all_rcWorks = this.RCWorks.Where(w => w.Number.StartsWith(msg_work.Number + "."));
                    foreach (RCWork rc_work in msg_work_all_rcWorks)
                    {
                        if (rc_work.ReportCard != null)
                        {
                            foreach (WorkDay rc_w_day in rc_work.ReportCard)
                            {
                                rc_w_day.LaborСosts = rc_w_day.Quantity * rc_work.Laboriousness;
                                if (msg_work.ReportCard == null)
                                {
                                    msg_work.ReportCard = new WorkReportCard();
                                    msg_work.ReportCard.Number = msg_work.Number;
                                }
                                WorkDay msg_w_day = msg_work.ReportCard.FirstOrDefault(wd => wd.Date == rc_w_day.Date);
                                if (msg_w_day == null)
                                {
                                    msg_w_day = new WorkDay();
                                    msg_w_day.Date = rc_w_day.Date;
                                    msg_w_day.LaborСosts += rc_w_day.LaborСosts;

                                }
                                else
                                    msg_w_day.LaborСosts += rc_w_day.LaborСosts;
                                if (msg_work.Laboriousness != 0)
                                    msg_w_day.Quantity = msg_w_day.LaborСosts / msg_work.Laboriousness;
                                msg_work.ReportCard.Add(msg_w_day);
                            }
                        }

                    }


                }
            }
        }
        public void CalcQuantity_backup()
        {
            List<MSGExellModel> loaded_models = new List<MSGExellModel>();
            foreach (WorksSection section in this.WorksSections)
            {
                foreach (MSGWork msg_work in section.MSGWorks)
                {
                    msg_work.Quantity = 0;
                    decimal common_vovr_labour_quantity = 0;
                    decimal common_vovr_previos_complate_labour_quantity = 0;
                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        vovr_work.Quantity = 0;
                        decimal common_ks_labour_quantity = 0;
                        decimal common_ks_previos_complate_labour_quantity = 0;
                        foreach (KSWork ks_work in vovr_work.KSWorks)
                        {
                            decimal common_rc_labour_quantity = 0;
                            decimal common_rc_previos_complate_labour_quantity = 0;
                            foreach (RCWork rc_work in ks_work.RCWorks)
                            {
                                rc_work.Quantity = 0;
                                if (this.Owner != null && rc_work.ReportCard != null)
                                {
                                    rc_work.Quantity = rc_work.ReportCard.Quantity + rc_work.ReportCard.PreviousComplatedQuantity;

                                }
                                else
                                {
                                    rc_work.PreviousComplatedQuantity = 0;

                                    rc_work.ReportCard = new WorkReportCard();
                                    this.Register(rc_work.ReportCard, "Number", rc_work.CellAddressesMap["Number"].Row, WRC_NUMBER_COL, this.RegisterSheet);
                                    this.Register(rc_work.ReportCard, "PreviousComplatedQuantity", rc_work.CellAddressesMap["Number"].Row, WRC_NUMBER_COL, this.RegisterSheet);
                                    rc_work.ReportCard.Number = rc_work.Number;
                                    int rc_work_common_quantity = 0;
                                    foreach (MSGExellModel model in this.Children)
                                    {
                                        if (!loaded_models.Contains(model))
                                        {
                                            model.LoadWorksReportCards();
                                            loaded_models.Add(model);
                                        }
                                        // RCWork child_rc_work = model.RCWorks.FirstOrDefault(w => w.Number == rc_work.Number);
                                        WorkReportCard child_rc = model.WorkReportCards.FirstOrDefault(rc => rc.Number == rc_work.Number);


                                        if (child_rc != null)
                                        {

                                            foreach (WorkDay child_w_day in child_rc)
                                            {
                                                WorkDay curent_w_day = rc_work.ReportCard.FirstOrDefault(wd => wd.Date == child_w_day.Date);
                                                if (curent_w_day != null)
                                                {
                                                    curent_w_day.Quantity += child_w_day.Quantity;
                                                    curent_w_day.LaborСosts = curent_w_day.Quantity * rc_work.Laboriousness;
                                                }
                                                else
                                                {
                                                    curent_w_day = new WorkDay();
                                                    curent_w_day.Date = child_w_day.Date;
                                                    curent_w_day.Quantity = child_w_day.Quantity;
                                                    curent_w_day.LaborСosts = child_w_day.Quantity * rc_work.Laboriousness;
                                                    DateTime end_date = DateTime.Parse(this.RegisterSheet.Cells[WORKS_END_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());

                                                    foreach (KeyValuePair<string, ExellPropAddress> map_item in child_w_day.CellAddressesMap)
                                                    {
                                                        int date_index = 0;
                                                        while (this.RegisterSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value != null &&
                                                                   DateTime.Parse(this.RegisterSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString()) < end_date)
                                                        {
                                                            if (DateTime.Parse(this.RegisterSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString()) == curent_w_day.Date)
                                                                break;
                                                            date_index++;
                                                        }
                                                        if (rc_work.ReportCard.CellAddressesMap.ContainsKey("Number"))
                                                        {
                                                            int curent_wrc_row = rc_work.ReportCard.CellAddressesMap["Number"].Row;
                                                            this.Register(curent_w_day, "Quantity", curent_wrc_row, WRC_DATE_COL + date_index, this.RegisterSheet);
                                                            curent_w_day.Quantity = curent_w_day.Quantity;
                                                        }

                                                    }
                                                    rc_work.ReportCard.Add(curent_w_day);
                                                }
                                            }
                                            rc_work.PreviousComplatedQuantity += child_rc.PreviousComplatedQuantity;
                                            rc_work.Quantity += child_rc.Quantity;
                                        }

                                    }
                                }
                                common_rc_labour_quantity += rc_work.Quantity * rc_work.Laboriousness;
                                common_rc_previos_complate_labour_quantity += rc_work.PreviousComplatedQuantity * rc_work.Laboriousness;
                            }

                            ks_work.Quantity = common_rc_labour_quantity / ks_work.Laboriousness;
                            ks_work.PreviousComplatedQuantity = common_rc_previos_complate_labour_quantity / ks_work.Laboriousness;

                            if (ks_work.Laboriousness != 0)
                            {
                                common_ks_labour_quantity += ks_work.Quantity * ks_work.Laboriousness;
                                common_ks_previos_complate_labour_quantity += ks_work.PreviousComplatedQuantity * ks_work.Laboriousness;
                            }
                        }

                        if (vovr_work.Laboriousness != 0)
                        {
                            vovr_work.Quantity = common_ks_labour_quantity / vovr_work.Laboriousness;
                            vovr_work.PreviousComplatedQuantity = common_ks_previos_complate_labour_quantity / vovr_work.Laboriousness;
                        }
                        common_vovr_labour_quantity += vovr_work.Quantity * vovr_work.Laboriousness;
                        common_vovr_previos_complate_labour_quantity += vovr_work.PreviousComplatedQuantity * vovr_work.Laboriousness;
                    }

                    if (msg_work.Laboriousness != 0)
                    {
                        msg_work.Quantity = common_vovr_labour_quantity / msg_work.Laboriousness;
                        msg_work.PreviousComplatedQuantity = common_vovr_previos_complate_labour_quantity / msg_work.Laboriousness;
                    }

                    var msg_work_all_rcWorks = this.RCWorks.Where(w => w.Number.StartsWith(msg_work.Number + "."));
                    foreach (RCWork rc_work in msg_work_all_rcWorks)
                    {
                        if (rc_work.ReportCard != null)
                        {
                            foreach (WorkDay ks_w_day in rc_work.ReportCard)
                            {
                                if (msg_work.ReportCard == null) msg_work.ReportCard = new WorkReportCard();
                                WorkDay msg_w_day = msg_work.ReportCard.FirstOrDefault(wd => wd.Date == ks_w_day.Date);
                                if (msg_w_day == null)
                                {
                                    msg_w_day = new WorkDay();
                                    msg_w_day.Date = ks_w_day.Date;
                                    msg_w_day.LaborСosts += ks_w_day.LaborСosts;

                                }
                                else
                                    msg_w_day.LaborСosts += ks_w_day.LaborСosts;
                                if (msg_work.Laboriousness != 0)
                                    msg_w_day.Quantity = msg_w_day.LaborСosts / msg_work.Laboriousness;
                                msg_work.ReportCard.Add(msg_w_day);
                            }
                        }

                    }


                }
            }
        }
        /// <summary>
        /// Функция подсчитывает потребления работчей силы для МСГ работ
        /// </summary>
        public void CalcWorkerConsumptions()
        {
            Excel.Worksheet consumtionsSheet = this.WorkerConsumptionsSheet;
            int rowIndex = W_CONSUMPTIONS_FIRST_ROW_INDEX;
            // this.WorkerConsumptions.Clear();
            null_str_count = 0;


            if (this.Owner == null)
                foreach (WorkerConsumption worker_consumption in this.WorkerConsumptions)
                {

                    foreach (MSGExellModel model in this.Children)
                    {

                        WorkerConsumption child_w_consumption = model.WorkerConsumptions.FirstOrDefault(w => w.Number == worker_consumption.Number);
                        if (child_w_consumption != null)
                        {

                            foreach (WorkerConsumptionDay child_w_day in child_w_consumption.WorkersConsumptionReportCard)
                            {
                                WorkerConsumptionDay curent_w_day = worker_consumption.WorkersConsumptionReportCard.FirstOrDefault(wd => wd.Date == child_w_day.Date);
                                if (curent_w_day != null)
                                {
                                    curent_w_day.Quantity += child_w_day.Quantity;
                                }
                                else
                                {
                                    curent_w_day = new WorkerConsumptionDay();
                                    //  this.Register(curent_w_day);
                                    curent_w_day.Date = child_w_day.Date;

                                    curent_w_day.Quantity = child_w_day.Quantity;

                                    DateTime end_date = DateTime.Parse(this.RegisterSheet.Cells[WORKS_END_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());

                                    foreach (KeyValuePair<string, ExellPropAddress> map_item in child_w_day.CellAddressesMap)
                                    {
                                        int date_index = 0;
                                        while (this.WorkerConsumptionsSheet.Cells[W_CONSUMPTIONS_DATE_RAW, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value != null &&
                                          DateTime.Parse(this.WorkerConsumptionsSheet.Cells[W_CONSUMPTIONS_DATE_RAW, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString()) < end_date)
                                        {
                                            if (DateTime.Parse(this.WorkerConsumptionsSheet.Cells[W_CONSUMPTIONS_DATE_RAW, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString()) == curent_w_day.Date)
                                                break;
                                            date_index++;
                                        }
                                        int curent_w_consumption_row = worker_consumption.WorkersConsumptionReportCard.CellAddressesMap["WorkersConsumptionReportCard"].Row;

                                        curent_w_day.CellAddressesMap.Add(map_item.Key, new ExellPropAddress(curent_w_consumption_row, W_CONSUMPTIONS_FIRST_DATE_COL + date_index, this.WorkerConsumptionsSheet));
                                        curent_w_day.Quantity = curent_w_day.Quantity;
                                        //          this.WorkerConsumptionsSheet.Cells[curent_w_consumption_row, W_CONSUMPTIONS_FIRST_DATE_COL + date_index] =
                                        //      curent_w_day.Quantity.ToString();

                                        //   this.Register(curent_w_day);
                                    }
                                    worker_consumption.WorkersConsumptionReportCard.Add(curent_w_day);
                                }
                            }
                        }

                    }
                }
        }
   
        /// <summary>
        /// Вычисление всех вычисляемых величин внутри модели и всех его дочерних моделей.
        /// </summary>
        public void CalcAll()
        {
            this.CalcLabourness();
            this.CalcQuantity();
            //  this.SetStyleFormats();

        }

        /// <summary>
        /// 
        /// или если ведомость сама общая, то просто очищает у нее каледарную часть с записями выполенных объемов
        /// </summary>
        public void UpdateExcelRepresetation()
        {
            this.UpdateCellAddressMapsWorkSheets();
            this.ClearWorksheetCommonPart();

            int last_row = FIRST_ROW_INDEX;
            foreach (WorksSection w_section in this.WorksSections.OrderBy(s => s.Number))
            {
                last_row = this.AdjustExcelRepresentionTree(w_section, last_row) + _SECTIONS_GAP;
                this.UpdateRepresentation(w_section);
            }

        }
        public void UpdateExcelRepresetation(IExcelBindableBase obj0)
        {

        }

        /// <summary>
        /// Функция копирует часть объектой модели из родительской модеи в текущую
        /// </summary>
        public void CopyOwnerObjectModels()
        {
            if (this.Owner != null)
            {
                this.Unregister(this.WorksSections);

                this.WorksSections = (ExcelNotifyChangedCollection<WorksSection>)this.Owner.WorksSections.Clone();
                this.WorksSections.Owner = this;
                this.SetCommonModelCollections();

                this.UpdateCellAddressMapsWorkSheets();
                foreach (var section in this.WorksSections)
                    this.RegisterObjectInObjectPropertyNameRegister(section);
            }

        }
        /// <summary>
        /// Функция заполняет соосветврующие общие коллекции из дерева загруженных в объекты данных
        /// </summary>
        public void SetCommonModelCollections()
        {
            this.MSGWorks.Clear();
            this.VOVRWorks.Clear();
            this.KSWorks.Clear();
            this.RCWorks.Clear();
            foreach (WorksSection w_section in this.WorksSections)
            {
                foreach (MSGWork msg_work in w_section.MSGWorks)
                {
                    if (!this.MSGWorks.Contains(msg_work)) this.MSGWorks.Add(msg_work);
                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        if (!this.VOVRWorks.Contains(vovr_work)) this.VOVRWorks.Add(vovr_work);

                        foreach (KSWork ks_work in vovr_work.KSWorks)
                        {
                            if (!this.KSWorks.Contains(ks_work)) this.KSWorks.Add(ks_work);
                            foreach (RCWork rc_work in ks_work.RCWorks)
                                if (!this.RCWorks.Contains(rc_work)) this.RCWorks.Add(rc_work);
                        }
                    }
                }

            }
        }
        /// <summary>
        /// 
        /// </summary>
        public void UpdateSectionExcelRepresentation_bkup()
        {
            foreach (WorksSection section in this.WorksSections)
                this.UpdateExellBindableObject(section);

            foreach (MSGWork msg_work in this.MSGWorks)
            {
                this.UpdateExellBindableObject(msg_work);

                foreach (WorkScheduleChunk w_ch in msg_work.WorkSchedules)
                    this.UpdateExellBindableObject(w_ch);
                foreach (NeedsOfWorker n_w in msg_work.WorkersComposition)
                    this.UpdateExellBindableObject(n_w);
            }
            foreach (VOVRWork vovr_work in this.VOVRWorks)
                this.UpdateExellBindableObject(vovr_work);

            foreach (KSWork ks_work in this.KSWorks)
                this.UpdateExellBindableObject(ks_work);

            foreach (RCWork rc_work in this.RCWorks)

                this.UpdateExellBindableObject(rc_work);
            foreach (var rc in this.WorkReportCards)
            {
                this.UpdateExellBindableObject(rc);
                foreach (WorkDay w_day in rc)
                    this.UpdateExellBindableObject(w_day);
            }


        }
        /// <summary>
        /// Функция отображает данные их модели в на соответвующие листа Excel
        /// </summary>
        /// <param name="w_section"></param>
        public void UpdateRepresentation(WorksSection w_section)
        {
            this.UpdateExellBindableObject(w_section);
            foreach (MSGWork msg_work in w_section.MSGWorks.OrderBy(w => w.Number))
            {
                this.UpdateRepresentation(msg_work);
            }
        }
        public void UpdateRepresentation(MSGWork msg_work)
        {
            this.UpdateExellBindableObject(msg_work);
            foreach (WorkScheduleChunk w_ch in msg_work.WorkSchedules)
                this.UpdateExellBindableObject(w_ch);
            foreach (NeedsOfWorker n_w in msg_work.WorkersComposition)
                this.UpdateExellBindableObject(n_w);
            foreach (NeedsOfMachine n_m in msg_work.MachinesComposition)
                this.UpdateExellBindableObject(n_m);
            foreach (VOVRWork vovr_work in msg_work.VOVRWorks.OrderBy(w => w.Number))
            {
                this.UpdateRepresentation(vovr_work);
            }
        }
        public void UpdateRepresentation(VOVRWork vovr_work)
        {
            this.UpdateExellBindableObject(vovr_work);
            foreach (KSWork ks_work in vovr_work.KSWorks.OrderBy(w => w.Number))
            {
                this.UpdateRepresentation(ks_work);
            }
        }
        public void UpdateRepresentation(KSWork ks_work)
        {
            this.UpdateExellBindableObject(ks_work);
            foreach (RCWork rc_work in ks_work.RCWorks.OrderBy(w => w.Number))
            {
                this.UpdateRepresentation(rc_work);
            }
        }
        public void UpdateRepresentation(RCWork rc_work)
        {
            this.UpdateExellBindableObject(rc_work);
            var rc_cards = this.WorkReportCards.Where(rc => rc.Number == rc_work.Number).ToList();

            foreach (var rc in rc_cards)
            {
                this.UpdateExellBindableObject(rc);
                foreach (WorkDay w_day in rc)
                    this.UpdateExellBindableObject(w_day);
            }
        }

        /// <summary>
        /// Функция устанавливает соовествующие значения строк и столбцов, где будут располагаться объеты на листе Excel
        /// </summary>
        /// <param name="w_section"></param>
        /// <param name="top_row"></param>
        /// <returns></returns>




        public int AdjustExcelRepresentionTree(WorksSection w_section, int top_row)
        {
            int section_row = top_row;
            int rc_row = top_row;
            int ks_row = top_row;
            int vovr_row = top_row;
            int msg_row = top_row;

            w_section.ChangeTopRow(section_row);
            foreach (MSGWork msg_work in w_section.MSGWorks.OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberSuffix}.", ""))))
            {
                msg_row = this.AdjustExcelRepresentionTree(msg_work, msg_row);
            }
            section_row = msg_row + 1;
            return rc_row;
        }
        
        public int AdjustExcelRepresentionTree(MSGWork msg_work, int msg_row)
        {
            int msg_lowest_row = 0;
            msg_work.ChangeTopRow(msg_row);
            int sh_ch_row_iterator = 0;
            foreach (WorkScheduleChunk w_ch in msg_work.WorkSchedules)
            {
                w_ch.ChangeTopRow(msg_row + sh_ch_row_iterator);
                sh_ch_row_iterator++;
            }
            int nw_row_iterator = 0;
            foreach (NeedsOfWorker n_w in msg_work.WorkersComposition)
            {
                n_w.ChangeTopRow(msg_row + nw_row_iterator);
                nw_row_iterator++;
            }

            var duple_msg_works = this.MSGWorks.Where(msgw => msgw.Number == msg_work.Number && msgw.Id != msg_work.Id).ToList();
            int msg_work_cuont = 0;
            foreach (var msgw in duple_msg_works)
            {
                msg_work_cuont++;
                msgw.ChangeTopRow(msg_row + msg_work_cuont);
            }
            if (msg_row + sh_ch_row_iterator > msg_lowest_row) msg_lowest_row = msg_row + sh_ch_row_iterator;
            if (msg_row + nw_row_iterator > msg_lowest_row) msg_lowest_row = msg_row + nw_row_iterator;
            if (msg_row + msg_work_cuont > msg_lowest_row) msg_lowest_row = msg_row + msg_work_cuont;

            msg_row += msg_work_cuont;
            int vovr_row = msg_row;
            foreach (VOVRWork vovr_work in msg_work.VOVRWorks.OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberSuffix}.", ""))))
            {


                vovr_row = this.AdjustExcelRepresentionTree(vovr_work, vovr_row); ;
            }
            if (vovr_row < msg_lowest_row)
                msg_row = msg_lowest_row + 1;
            else
                msg_row = vovr_row + 1;

            return msg_row;
        }
        public int AdjustExcelRepresentionTree(VOVRWork vovr_work, int vovr_row)
        {
            vovr_work.ChangeTopRow(vovr_row);
            var duple_vovr_works = this.VOVRWorks.Where(vrw => vrw.Number == vovr_work.Number && vrw.Id != vovr_work.Id).ToList();
            int vovr_work_cuont = 0;
            foreach (var vrw in duple_vovr_works)
            {
                vovr_work_cuont++;
                vrw.ChangeTopRow(vovr_row + vovr_work_cuont);
            }
            int ks_row = vovr_row;
            foreach (KSWork ks_work in vovr_work.KSWorks.OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberSuffix}.", ""))))
            {
                ks_row = this.AdjustExcelRepresentionTree(ks_work, ks_row); ;
            }

            vovr_row = ks_row;
            return vovr_row;
        }

        public int AdjustExcelRepresentionTree(KSWork ks_work, int ks_row)
        {
            ks_work.ChangeTopRow(ks_row);
            var duple_kc_works = this.KSWorks.Where(ksw => ksw.Number == ks_work.Number && ksw.Id != ks_work.Id).ToList();
            int ks_work_cuont = 0;
            foreach (var ksw in duple_kc_works)
            {
                ks_work_cuont++;
                ksw.ChangeTopRow(ks_row + ks_work_cuont);
            }

            int rc_row = ks_row + ks_work_cuont;
            foreach (RCWork rc_work in ks_work.RCWorks.OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberSuffix}.", ""))))
            {
                rc_row = this.AdjustExcelRepresentionTree(rc_work, rc_row);
                rc_row++;
            }

            ks_row = rc_row;
            return ks_row;
        }

        public int AdjustExcelRepresentionTree(RCWork rc_work, int rc_row)
        {
            rc_work.ChangeTopRow(rc_row);
            ///Находимо работы с таким же номером и помещаем их ниже 
            var duple_rc_works = this.RCWorks.Where(rcw => rcw.Number == rc_work.Number && rcw.Id != rc_work.Id).ToList();
            int rc_work_cuont = 0;
            foreach (var rcw in duple_rc_works)
            {
                rc_work_cuont++;
                rcw.ChangeTopRow(rc_row + rc_work_cuont);
            }

            if (rc_work.ReportCard != null)
            {
                rc_work.ReportCard.ChangeTopRow(rc_row);
                var duple_rc_work_rc = this.WorkReportCards.Where(rc => rc.Number == rc_work.Number && rc.Id != rc_work.ReportCard.Id).ToList();
                int rc_card_count = 0;
                foreach (WorkReportCard rc in duple_rc_work_rc)
                {
                    rc_card_count++;
                    rc.ChangeTopRow(rc_row + rc_card_count);
                    foreach (WorkDay w_day in rc)
                    {
                        w_day.ChangeTopRow(rc_work.CellAddressesMap["Number"].Row);
                    }
                }

                if (rc_work_cuont > rc_card_count)
                    rc_row += rc_work_cuont;
                else
                    rc_row += rc_card_count;
            }
            return rc_row;
        }

       

        public void Update()
        {
            this.UpdateExcelRepresetation();
        }
        /// <summary>
        /// Функция устанавливаетв объектах текущией модели соотвествующие worksheet-ы 
        /// Применяется в основном после применения Clone()  к объектоной модели MSGExcellModel
        /// </summary>
        public void UpdateCellAddressMapsWorkSheets()
        {

            this.WorksSections.CellAddressesMap.SetWorksheet(this.RegisterSheet);
            foreach (WorksSection w_section in this.WorksSections)
            {
                w_section.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                foreach (MSGWork msg_work in w_section.MSGWorks)
                {
                    msg_work.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                    //   if (!this.MSGWorks.Contains(msg_work)) this.MSGWorks.Add(msg_work);

                    foreach (WorkScheduleChunk w_ch in msg_work.WorkSchedules)
                        w_ch.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                    foreach (NeedsOfWorker n_w in msg_work.WorkersComposition)
                        n_w.CellAddressesMap.SetWorksheet(this.RegisterSheet);

                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        //  if (!this.VOVRWorks.Contains(vovr_work)) this.VOVRWorks.Add(vovr_work);
                        vovr_work.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                        foreach (KSWork ks_work in vovr_work.KSWorks)
                        {
                            //      if (!this.KSWorks.Contains(ks_work)) this.KSWorks.Add(ks_work);
                            ks_work.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                            foreach (RCWork rc_work in ks_work.RCWorks)
                            {
                                //    if (!this.RCWorks.Contains(rc_work)) this.RCWorks.Add(rc_work);
                                rc_work.ReportCard = this.WorkReportCards.Where(rc => rc.Number == rc_work.Number).FirstOrDefault();
                                rc_work.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                            }
                        }
                    }
                }

            }
        }
        /// <summary>
        /// Фунция очищает календарную часть ведомости (очищает все записи выполненных работ)
        /// </summary>
        public void ClearWorksheetDaysPart()
        {
            //    Excel.Range last_cell = this.RegisterSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            Excel.Range common_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, WRC_DATE_COL],
                this.RegisterSheet.Cells[this.KSWorks[this.KSWorks.Count - 1].CellAddressesMap["Laboriousness"].Row,
                                          this.KSWorks[this.KSWorks.Count - 1].CellAddressesMap["Laboriousness"].Column]];
            if (this.Owner != null)
                common_area_range.ClearContents();

            common_area_range.Interior.ColorIndex = 0;

            //  last_cell = this.WorkerConsumptionsSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            var last_work_consumption_RC = this.WorkerConsumptions[this.WorkerConsumptions.Count - 1].WorkersConsumptionReportCard;
            WorkerConsumptionDay last_work_consumption_Day = null;
            foreach (WorkerConsumption worker_consumption in this.WorkerConsumptions)
            {
                foreach (WorkerConsumptionDay worker_consumption_day in worker_consumption.WorkersConsumptionReportCard)
                {
                    if (last_work_consumption_Day == null || last_work_consumption_Day.Date < worker_consumption_day.Date)
                        last_work_consumption_Day = worker_consumption_day;
                }
            }

            if (last_work_consumption_Day != null)
            {
                common_area_range = this.WorkerConsumptionsSheet.Range[this.WorkerConsumptionsSheet.Cells[W_CONSUMPTIONS_FIRST_ROW_INDEX, W_CONSUMPTIONS_FIRST_DATE_COL],
                            this.WorkerConsumptionsSheet.Cells[last_work_consumption_Day.CellAddressesMap["Quantity"].Row, last_work_consumption_Day.CellAddressesMap["Quantity"].Column]];
                common_area_range.ClearContents();
                common_area_range.Interior.ColorIndex = 0;
            }


        }
        /// <summary>
        /// Функия очищает левую часть вдомости с МСГ, ВОВР и КС-2.
        /// </summary>
        public void ClearWorksheetCommonPart()
        {
            try
            {
                Excel.Range last_cell = this.RegisterSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

                Excel.Range common_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, WSEC_NUMBER_COL],
                      this.RegisterSheet.Cells[last_cell.Row, WRC_NUMBER_COL - 1]];
                common_area_range.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                common_area_range.Interior.ColorIndex = 0;

                Excel.Range record_cards_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, RC_NUMBER_COL],
                      this.RegisterSheet.Cells[last_cell.Row, last_cell.Column]];
                record_cards_area_range.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                record_cards_area_range.Interior.ColorIndex = 0;

                common_area_range.ClearContents();
                record_cards_area_range.ClearContents();
            }
            catch
            {

            }

            this.RemoveGroups(this.RegisterSheet);

        }
        /// <summary>
        /// Удалфет все групы в сторках и столбцах
        /// </summary>
        /// <param name="worksheet"></param>
        public void RemoveGroups(Excel.Worksheet worksheet)
        {
            Excel.Range all_rows = worksheet.Cells.Rows;
            Excel.Range all_colomns = worksheet.Cells.Columns;


            for (int ii = 0; ii < 5; ii++)
                try
                {
                    all_rows.Select();
                    all_rows.Ungroup();
                    all_colomns.Select();
                    all_colomns.Ungroup();
                }
                catch
                {
                }
        }
        /// <summary>
        /// Функция получает ближайший на листе Exсуд объет необходимого типа.
        /// </summary>
        /// <param name="section"></param>
        /// <param name="object_type"></param>
        /// <returns></returns>
        public IExcelBindableBase GetObjectBySelection(Excel.Range section, Type object_type)
        {
            ObservableCollection<Tuple<double, IExcelBindableBase>> objects_distation = new ObservableCollection<Tuple<double, IExcelBindableBase>>();

            foreach (var kvp in this.RegistedObjects.Where(rr => rr.Entity.GetType() == object_type))
            {
                int obj_row = kvp.ExellPropAddress.Row;
                int obj_col = kvp.ExellPropAddress.Column;
                double dist = Math.Sqrt(Math.Pow(obj_row - section.Row, 2) + Math.Pow(obj_col - section.Column, 2));

                objects_distation.Add(new Tuple<double, IExcelBindableBase>(dist, kvp.Entity));
            }
            IExcelBindableBase finded_obj = null;
            var tuple = objects_distation.OrderBy(el => el.Item1).FirstOrDefault();

            if (tuple != null)
                finded_obj = tuple.Item2 as IExcelBindableBase;

            return finded_obj;
        }
    }
}
