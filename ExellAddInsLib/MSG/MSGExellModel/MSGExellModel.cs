
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using static System.Collections.Specialized.BitVector32;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public partial class MSGExellModel : ExellModelBase
    {
        public const int SECTIONS_NULL_COUNTER = 500;
        public const int WORKS_NULL_COUNTER = 100;


        public const int COMMON_PARAMETRS_VALUE_COL = 3; //Номер стобца с общим параметрами проекта

        public const int CONTRACT_CODE_ROW = 2; //Код объекта или договора
        public const int CONSTRUCTION_OBJECT_CODE_ROW = 3;// Код объекта
        public const int CONSTRUCTION_SUBOBJECT_CODE_ROW = 4;//Код подъобьекта

        public const int WORKS_START_DATE_ROW = 1;
        public const int WORKS_START_DATE_COL = 3;
        public const int WORKS_END_DATE_ROW = 2;
        public const int WORKS_END_DATE_COL = 3;
        public const int FIRST_ROW_INDEX = 8;

        public const int _SECTIONS_GAP = 2;
        public const int _MSG_WORKS_GAP = 1;

        public const int _SECTIONS_GAP_FOR_INVALID_OBJECTS = 5;


        public const int W_SECTION_COLOR = 33;


        //public Guid LaodSessionId { get; set; }

        private int null_str_count = 0;

        private DateTime _worksStartDate;
        /// <summary>
        /// Дата начала ведомости 
        /// </summary>
        public DateTime WorksStartDate
        {
            get
            {
                return _worksStartDate;
            }
            set
            {
                SetProperty(ref _worksStartDate, value);
            }
        }
        //public DateTime WorksStartDate
        //{
        //    get
        //    {
        //        DateTime end_date = DateTime.MinValue;
        //        var last_ended_work = this.MSGWorks.OrderBy(w => w.WorkSchedules.StartDate).FirstOrDefault();
        //        if (last_ended_work != null)
        //            end_date = last_ended_work.WorkSchedules.StartDate;
        //        return end_date;
        //    }

        //}
        private DateTime _recordCardStartDate;

        public DateTime RecordCardStartDate
        {
            get { return _recordCardStartDate; }
            set { SetProperty(ref _recordCardStartDate, value); }
        }

        private DateTime _worksEndDate;
        /// <summary>
        /// Дата окончания работ в данной ведомости в соотвествии с планируемыми в  сроками отраженнным в части МСГ ведомости.
        /// (в части WorkSchedules работ MSGWork)
        /// </summary>
        public DateTime WorksEndDate
        {
            get
            {
                DateTime end_date = _worksEndDate;
                var last_ended_work = this.MSGWorks.OrderBy(w => w.WorkSchedules.EndDate).LastOrDefault();
                if (last_ended_work != null)
                    end_date = last_ended_work.WorkSchedules.EndDate;

                return end_date;
            }
            set
            {
                _worksEndDate = value;
            }
        }

        private int _workedDaysNumber;
        /// <summary>
        /// Общее количество каленадрых дне с начала до окончания работ
        /// </summary>
        public int WorkedDaysNumber
        {
            get
            {
                //   if (this.WorksEndDate == DateTime.MinValue)
                //     return _workedDaysNumber;
                return (this.WorksEndDate - this.WorksStartDate).Days;
            }

            set
            {
                this.WorksEndDate = this.WorksStartDate.AddDays(value);
                _workedDaysNumber = value;
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



        private ExcelNotifyChangedCollection<UnitOfMeasurement> _unitOfMeasurements;
        /// <summary>
        /// Коллекция с единицами измерения модели
        /// </summary>
        public ExcelNotifyChangedCollection<UnitOfMeasurement> UnitOfMeasurements
        {
            get { return _unitOfMeasurements; }
            set { SetProperty(ref _unitOfMeasurements, value); }
        }

        private ExcelNotifyChangedCollection<NeedsOfWorker> _WorkersComposition;
        /// <summary>
        /// Состав работников ( потребности)
        /// </summary>
        public ExcelNotifyChangedCollection<NeedsOfWorker> WorkersComposition
        {
            get { return _WorkersComposition; }
            set { SetProperty(ref _WorkersComposition, value); }
        }
        private ExcelNotifyChangedCollection<NeedsOfMachine> _machinesComposition;
        /// <summary>
        /// Состав работников ( потребности)
        /// </summary>
        public ExcelNotifyChangedCollection<NeedsOfMachine> MachinesComposition
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


        private ExcelNotifyChangedCollection<IObservableExcelBindableBase> _invalidObjects = new ExcelNotifyChangedCollection<IObservableExcelBindableBase>();
        /// <summary>
        /// Коллекция с единицами измерения модели
        /// </summary>
        public ExcelNotifyChangedCollection<IObservableExcelBindableBase> InvalidObjects
        {
            get { return _invalidObjects; }
            set { SetProperty(ref _invalidObjects, value); }
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
                this.WorksSections.Worksheet = _registerSheet;
                this.MSGWorks.Worksheet = _registerSheet;
                this.VOVRWorks.Worksheet = _registerSheet;
                this.KSWorks.Worksheet = _registerSheet;
                this.RCWorks.Worksheet = _registerSheet;
                this.WorkReportCards.Worksheet = _registerSheet;
                this.WorkersComposition.Worksheet = _registerSheet;
                this.MachinesComposition.Worksheet = _registerSheet;

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
                this.WorkerConsumptions.Worksheet = _workerConsumptionsSheet;
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
                this.MachineConsumptions.Worksheet = _machineConsumptionsSheet;
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
            WorksSections.Owner = this;
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
        public void LoadWorksSection(int row)
        {
            int rowIndex = row;
            Excel.Worksheet registerSheet = this.RegisterSheet;
            string number = registerSheet.Cells[rowIndex, WorksSection.WSEC_NUMBER_COL].Value.ToString();
            WorksSection w_section = this.WorksSections.FirstOrDefault(w => w.Number == number);
            if (w_section == null)
                w_section = new WorksSection();

            w_section.LoadSessionId = this.LoadSessionId;
            w_section.Worksheet = registerSheet;
            this.Register(w_section, "Number", rowIndex, WorksSection.WSEC_NUMBER_COL, registerSheet);
            this.Register(w_section, "Name", rowIndex, WorksSection.WSEC_NAME_COL, registerSheet);

            w_section.Number = number;
            // if (this.WorksSections.FirstOrDefault(ws => ws.Number == w_section.Number) != null)
            //   w_section.CellAddressesMap["Number"].IsValid = false;

            var name = registerSheet.Cells[rowIndex, WorksSection.WSEC_NAME_COL].Value;
            if (name != null)
                w_section.Name = name;
            else
                w_section.SetPropertyValidStatus("Name", false);

            if (!this.WorksSections.Contains(w_section))
            {
                this.WorksSections.Add(w_section);
                w_section.Owner = this;
            }
        }
        /// <summary>
        /// Функция из части РАЗДЕЛЫ  листа Worksheet создает и помещает в модель  разделы работ
        /// </summary>
        public void LoadWorksSections()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;

            while (null_str_count < SECTIONS_NULL_COUNTER)
            {
                var number = registerSheet.Cells[rowIndex, WorksSection.WSEC_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    this.LoadWorksSection(rowIndex);
                }

                rowIndex++;
            }
        }
        public void LoadMSGWork(int row, int dist_row = HASH_FUNCTION_ROW)
        {
            int rowIndex = row;
            Excel.Worksheet registerSheet = this.RegisterSheet;
            string number = registerSheet.Cells[rowIndex, MSGWork.MSG_NUMBER_COL].Value;
            if (number == null) return;
            MSGWork msg_work = this.MSGWorks.FirstOrDefault(w => w.Number == number);
            if (msg_work == null)
                msg_work = new MSGWork();

            msg_work.LoadSessionId = this.LoadSessionId;

            msg_work.Worksheet = registerSheet;
            this.Register(msg_work, "Number", rowIndex, MSGWork.MSG_NUMBER_COL, this.RegisterSheet, false, (v) => Regex.IsMatch((string)v, @"^\d+\.\d+$"));
            this.Register(msg_work, "Name", rowIndex, MSGWork.MSG_NAME_COL, this.RegisterSheet);
            this.Register(msg_work, "ProjectQuantity", rowIndex, MSGWork.MSG_QUANTITY_COL, this.RegisterSheet, false, (v) => ((decimal)v) != 0);
            this.Register(msg_work, "Quantity", rowIndex, MSGWork.MSG_QUANTITY_FACT_COL, this.RegisterSheet, true);
            this.Register(msg_work, "Laboriousness", rowIndex, MSGWork.MSG_LABOURNESS_COL, this.RegisterSheet);
            this.Register(msg_work, "UnitOfMeasurement.Name", rowIndex, MSGWork.MSG_MEASURE_COL, registerSheet);

            msg_work.Number = number;
            //  if (this.MSGWorks.FirstOrDefault(w => w.Number == msg_work.Number) != null && this.Owner == null)
            //     msg_work.CellAddressesMap["Number"].IsValid = false;

            var name = registerSheet.Cells[rowIndex, MSGWork.MSG_NAME_COL].Value;
            if (name != null)
                msg_work.Name = name;
            else
                msg_work.SetPropertyValidStatus("Name", false); ;

            var unit_of_measurement_name = registerSheet.Cells[rowIndex, MSGWork.MSG_MEASURE_COL].Value;
            if (unit_of_measurement_name != null)
                msg_work.UnitOfMeasurement = UnitOfMeasurements.FirstOrDefault(um => um.Name == unit_of_measurement_name.ToString());
            else
                msg_work.SetPropertyValidStatus("UnitOfMeasurement.Name", false);

            var pr_quantity = registerSheet.Cells[rowIndex, MSGWork.MSG_QUANTITY_COL].Value;
            if (pr_quantity != null)
                msg_work.ProjectQuantity = Decimal.Parse(pr_quantity.ToString());
            else
                msg_work.SetPropertyValidStatus("ProjectQuantity", false);

            var quantity = registerSheet.Cells[rowIndex, MSGWork.MSG_QUANTITY_FACT_COL].Value;
            if (quantity != null)
                msg_work.Quantity = Decimal.Parse(quantity.ToString());

            var labourness = registerSheet.Cells[rowIndex, MSGWork.MSG_LABOURNESS_COL].Value;
            if (labourness != null)
                msg_work.Laboriousness = Decimal.Parse(labourness.ToString());
            else
                msg_work.SetPropertyValidStatus("Laboriousness", false);

            this.LoadWorksSchedules(msg_work);
            this.LoadMSGWorkerCompositions(msg_work);
            this.LoadMSGMachineCompositions(msg_work);

            if (!this.MSGWorks.Contains(msg_work))
                this.MSGWorks.Add(msg_work);
        }
        public void LoadWorksSchedules(MSGWork msg_work)
        {
            Excel.Worksheet registerSheet = this.RegisterSheet; ;
            int rowIndex = msg_work.GetTopRow();
            if (msg_work.Number == "1.1")
                ;
            int schedule_number = 0;

            while (registerSheet.Cells[rowIndex, MSGWork.MSG_START_DATE_COL].Value != null)
            {
                schedule_number++;
                int duration = 0;
                DateTime start_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSGWork.MSG_START_DATE_COL].Value.ToString());
                Excel.Range end_date_cell = registerSheet.Cells[rowIndex, MSGWork.MSG_END_DATE_COL];
                Excel.Range duration_cell = registerSheet.Cells[rowIndex, MSGWork.MSG_DURATION_DATE_COL];
                DateTime end_time;

                if (start_time < this.WorksStartDate) this.WorksStartDate = start_time;

                if (duration_cell.Value != null && Int32.TryParse(duration_cell.Value.ToString(), out duration) && duration != 0)
                {
                    end_time = start_time.AddDays(duration - 1);
                    end_date_cell.Value = end_time.ToString("d");
                }
                else
                {
                    end_time = DateTime.Parse(end_date_cell.Value.ToString());
                    duration_cell.NumberFormat = "0";
                    duration_cell.Value = (end_time - start_time).Days + 1;
                }

                WorkScheduleChunk work_sh_chunk = msg_work.WorkSchedules.FirstOrDefault(shd => shd.Number == $"{msg_work.Number}.{schedule_number}");
                if (work_sh_chunk == null)
                    work_sh_chunk = new WorkScheduleChunk(start_time, end_time);
                else
                {
                    work_sh_chunk.StartTime = start_time;
                    work_sh_chunk.EndTime = end_time;
                }

                work_sh_chunk.Worksheet = registerSheet;
                work_sh_chunk.Number = $"{msg_work.Number}.{schedule_number.ToString()}";
                work_sh_chunk.Duration = duration;

                string is_snaday_vacation = registerSheet.Cells[rowIndex, MSGWork.MSG_SUNDAY_IS_VOCATION_COL].Value;
                is_snaday_vacation = registerSheet.Cells[rowIndex, MSGWork.MSG_SUNDAY_IS_VOCATION_COL].Value;
                if (is_snaday_vacation != null && is_snaday_vacation.Contains("Нет"))
                    work_sh_chunk.IsSundayVacationDay = "Нет";
                else
                    work_sh_chunk.IsSundayVacationDay = "Да";

                this.Register(work_sh_chunk, "StartTime", rowIndex, MSGWork.MSG_START_DATE_COL, this.RegisterSheet);
                this.Register(work_sh_chunk, "EndTime", rowIndex, MSGWork.MSG_END_DATE_COL, this.RegisterSheet);
                this.Register(work_sh_chunk, "Duration", rowIndex, MSGWork.MSG_DURATION_DATE_COL, this.RegisterSheet);
                this.Register(work_sh_chunk, "IsSundayVacationDay", rowIndex, MSGWork.MSG_SUNDAY_IS_VOCATION_COL, this.RegisterSheet);

                if (!msg_work.WorkSchedules.Contains(work_sh_chunk))
                {
                    if (msg_work.WorkSchedules.IsIntersections(work_sh_chunk) || start_time > end_time)
                    {
                        work_sh_chunk.IsValid = false;
                        work_sh_chunk.SetPropertyValidStatus("StartTime", false);
                        work_sh_chunk.SetPropertyValidStatus("EndTime", false);
                        if (this.Owner == null)
                            throw new Exception("Диапазона дат начали и конца МСГ работы пересекаются с уже имещимися!!");
                    }
                    else
                        msg_work.WorkSchedules.Add(work_sh_chunk);

                }
                work_sh_chunk.Owner = msg_work;
                if (registerSheet.Cells[rowIndex + 1, MSGWork.MSG_NUMBER_COL].Value != null) break;
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
            this.WorkersComposition.Clear();
            this.MachinesComposition.Clear();

            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, MSGWork.MSG_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    this.LoadMSGWork(rowIndex);
                }
                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из части МСГ листа Worksheet 
        /// </summary>
        public void LoadMSGWorkerCompositions(MSGWork msg_work)
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;

            int rowIndex = msg_work.GetTopRow();
            int need_number = 0;
            while (registerSheet.Cells[rowIndex, MSGWork.MSG_NEEDS_OF_WORKERS_NAME_COL].Value != null)
            {
                need_number++;
                NeedsOfWorker msg_needs_of_workers = msg_work.WorkersComposition.FirstOrDefault(n_w => n_w.Number == $"{msg_work.Number}.{need_number}");// n_w.NumberSuffix == msg_work.Number);
                if (msg_needs_of_workers == null)
                    msg_needs_of_workers = new NeedsOfWorker();
                msg_needs_of_workers.NeedsOfWorkersReportCard.Clear();

                msg_needs_of_workers.Worksheet = registerSheet;
                //   this.Register(msg_needs_of_workers, "Number", rowIndex,MSGWork.MSG_NEEDS_OF_WORKERS_NUMBER_COL, this.RegisterSheet);
                this.Register(msg_needs_of_workers, "Name", rowIndex, MSGWork.MSG_NEEDS_OF_WORKERS_NAME_COL, this.RegisterSheet);
                this.Register(msg_needs_of_workers, "Quantity", rowIndex, MSGWork.MSG_NEEDS_OF_WORKERS_QUANTITY_COL, this.RegisterSheet, true);

                msg_needs_of_workers.Number = $"{msg_work.Number}.{need_number.ToString()}";
                msg_needs_of_workers.Name = registerSheet.Cells[rowIndex, MSGWork.MSG_NEEDS_OF_WORKERS_NAME_COL].Value;

                var quantity = registerSheet.Cells[rowIndex, MSGWork.MSG_NEEDS_OF_WORKERS_QUANTITY_COL].Value;
                if (quantity != null)
                    msg_needs_of_workers.Quantity = Decimal.Parse(quantity.ToString());
                else
                    msg_needs_of_workers.SetPropertyValidStatus("Quantity", false);

                if (!msg_work.WorkersComposition.Contains(msg_needs_of_workers))
                    msg_work.WorkersComposition.Add(msg_needs_of_workers);

                msg_needs_of_workers.Owner = msg_work;
                foreach (WorkScheduleChunk chunk in msg_work.WorkSchedules)
                {
                    for (DateTime date = chunk.StartTime; date <= chunk.EndTime; date = date.AddDays(1))
                    {
                        if (!(msg_work.IsSundayVocation && date.DayOfWeek == DayOfWeek.Sunday))
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
                            //   if(!global_needs_of_worker.NeedsOfWorkersReportCard.Contains(new_nw_day))
                            global_needs_of_worker.NeedsOfWorkersReportCard.Add(new_nw_day);
                        }
                    }

                }
                if (!this.WorkersComposition.Contains(global_needs_of_worker))
                    this.WorkersComposition.Add(global_needs_of_worker);

                rowIndex++;
            }
        }

        /// <summary>
        /// Функция из части МСГ листа Worksheet 
        /// </summary>
        public void LoadMSGMachineCompositions(MSGWork msg_work)
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;

            int rowIndex = msg_work.GetTopRow();
            int need_number = 0;
            while (registerSheet.Cells[rowIndex, MSGWork.MSG_NEEDS_OF_MACHINE_NAME_COL].Value != null)
            {

                need_number++;
                NeedsOfMachine msg_needs_of_machines = msg_work.MachinesComposition.FirstOrDefault(n_m => n_m.Number == $"{msg_work.Number}.{need_number}"); //n_m.NumberSuffix == msg_work.Number);
                if (msg_needs_of_machines == null)
                    msg_needs_of_machines = new NeedsOfMachine();

                msg_needs_of_machines.NeedsOfMachinesReportCard.Clear();

                msg_needs_of_machines.Worksheet = registerSheet;
                //    this.Register(msg_needs_of_machines, "Number", rowIndex,MSGWork.MSG_NEEDS_OF_MACHINE_NUMBER_COL, this.RegisterSheet);
                this.Register(msg_needs_of_machines, "Name", rowIndex, MSGWork.MSG_NEEDS_OF_MACHINE_NAME_COL, this.RegisterSheet);
                this.Register(msg_needs_of_machines, "Quantity", rowIndex, MSGWork.MSG_NEEDS_OF_MACHINE_QUANTITY_COL, this.RegisterSheet, false);

                msg_needs_of_machines.Number = $"{msg_work.Number}.{need_number.ToString()}";
                msg_needs_of_machines.Name = registerSheet.Cells[rowIndex, MSGWork.MSG_NEEDS_OF_MACHINE_NAME_COL].Value;

                var quantity = registerSheet.Cells[rowIndex, MSGWork.MSG_NEEDS_OF_MACHINE_QUANTITY_COL].Value;
                if (quantity != null)
                    msg_needs_of_machines.Quantity = decimal.Parse(quantity.ToString());
                else
                    msg_needs_of_machines.SetPropertyValidStatus("Quantity", false);

                if (!msg_work.MachinesComposition.Contains(msg_needs_of_machines))
                    msg_work.MachinesComposition.Add(msg_needs_of_machines);

                msg_needs_of_machines.Owner = msg_work;
                foreach (WorkScheduleChunk chunk in msg_work.WorkSchedules)
                {
                    for (DateTime date = chunk.StartTime; date <= chunk.EndTime; date = date.AddDays(1))
                    {
                        if (!(msg_work.IsSundayVocation && date.DayOfWeek == DayOfWeek.Sunday))
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
                    if (!this.MachinesComposition.Contains(global_needs_of_machine))
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
                    if (!this.MachinesComposition.Contains(msg_needs_of_machines))
                        this.MachinesComposition.Add(msg_needs_of_machines);
                }

                rowIndex++;
            }
        }

        public void LoadVOVRWork(int row)
        {
            int rowIndex = row;
            Excel.Worksheet registerSheet = this.RegisterSheet;
            string number = registerSheet.Cells[rowIndex, VOVRWork.VOVR_NUMBER_COL].Value;
            if (number == null) return;
            VOVRWork vovr_work = this.VOVRWorks.FirstOrDefault(w => w.Number == number);
            if (vovr_work == null)
                vovr_work = new VOVRWork();

            vovr_work.LoadSessionId = this.LoadSessionId;
            vovr_work.Worksheet = registerSheet;
            this.Register(vovr_work, "Number", rowIndex, VOVRWork.VOVR_NUMBER_COL, this.RegisterSheet, false, v => Regex.IsMatch(v.ToString(), @"^\d+\.\d+\.\d+$"));
            this.Register(vovr_work, "Name", rowIndex, VOVRWork.VOVR_NAME_COL, this.RegisterSheet);
            this.Register(vovr_work, "ProjectQuantity", rowIndex, VOVRWork.VOVR_QUANTITY_COL, this.RegisterSheet, false, (v) => ((decimal)v) != 0);
            this.Register(vovr_work, "Quantity", rowIndex, VOVRWork.VOVR_QUANTITY_FACT_COL, this.RegisterSheet, true);
            this.Register(vovr_work, "Laboriousness", rowIndex, VOVRWork.VOVR_LABOURNESS_COL, this.RegisterSheet);
            this.Register(vovr_work, "UnitOfMeasurement.Name", rowIndex, VOVRWork.VOVR_MEASURE_COL, this.RegisterSheet);

            vovr_work.Number = number;

            //  if (this.VOVRWorks.FirstOrDefault(w => w.Number == vovr_work.Number) != null && this.Owner == null)
            //   vovr_work.CellAddressesMap["Number"].IsValid = false;

            var name = registerSheet.Cells[rowIndex, VOVRWork.VOVR_NAME_COL].Value;
            if (name != null)
                vovr_work.Name = name.ToString();
            else
                vovr_work.SetPropertyValidStatus("Name", false);
            var unit_of_measurement_name = registerSheet.Cells[rowIndex, VOVRWork.VOVR_MEASURE_COL].Value;
            if (unit_of_measurement_name != null)
                vovr_work.UnitOfMeasurement = new UnitOfMeasurement(unit_of_measurement_name.ToString());
            else
                vovr_work.SetPropertyValidStatus("UnitOfMeasurement.Name", false);

            var pr_quantity = registerSheet.Cells[rowIndex, VOVRWork.VOVR_QUANTITY_COL].Value;
            if (pr_quantity != null)
                vovr_work.ProjectQuantity = Decimal.Parse(pr_quantity.ToString());
            else
                vovr_work.SetPropertyValidStatus("ProjectQuantity", false);

            var quantity = registerSheet.Cells[rowIndex, VOVRWork.VOVR_QUANTITY_FACT_COL].Value;
            if (quantity != null)
                vovr_work.Quantity = Decimal.Parse(quantity.ToString());



            var labouriosness = registerSheet.Cells[rowIndex, VOVRWork.VOVR_LABOURNESS_COL].Value;
            if (labouriosness != null)
                vovr_work.Laboriousness = Decimal.Parse(labouriosness.ToString());
            else
                vovr_work.SetPropertyValidStatus("Laboriousness", false);

            if (!this.VOVRWorks.Contains(vovr_work))
                this.VOVRWorks.Add(vovr_work);

        }
        /// <summary>
        /// Функция из части  ВОВР листа Worksheet создает и помещает в модель работы типа VOVRWork 
        /// </summary>
        public void LoadVOVRWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;


            null_str_count = 0;
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, VOVRWork.VOVR_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    this.LoadVOVRWork(rowIndex);
                }

                rowIndex++;
            }
        }
        public void LoadKSWork(int row)
        {
            int rowIndex = row;
            Excel.Worksheet registerSheet = this.RegisterSheet;
            string number = registerSheet.Cells[rowIndex, KSWork.KS_NUMBER_COL].Value;
            if (number == null) return;
            KSWork ks_work = this.KSWorks.FirstOrDefault(w => w.Number == number);
            if (ks_work == null)
                ks_work = new KSWork();

            ks_work.LoadSessionId = this.LoadSessionId;
            ks_work.Worksheet = registerSheet;

            this.Register(ks_work, "Number", rowIndex, KSWork.KS_NUMBER_COL, this.RegisterSheet, false, v => Regex.IsMatch(v.ToString(), @"^\d+\.\d+\.\d+\.\d+$"));
            this.Register(ks_work, "EstimateNumber", rowIndex, KSWork.KS_ESTIMATE_NUMBER_COL, this.RegisterSheet);
            this.Register(ks_work, "Code", rowIndex, KSWork.KS_CODE_COL, this.RegisterSheet);
            this.Register(ks_work, "Name", rowIndex, KSWork.KS_NAME_COL, this.RegisterSheet);
            this.Register(ks_work, "ProjectQuantity", rowIndex, KSWork.KS_QUANTITY_COL, this.RegisterSheet);
            this.Register(ks_work, "Quantity", rowIndex, KSWork.KS_QUANTITY_FACT_COL, this.RegisterSheet, true);
            this.Register(ks_work, "Laboriousness", rowIndex, KSWork.KS_LABOURNESS_COL, this.RegisterSheet);
            this.Register(ks_work, "UnitOfMeasurement.Name", rowIndex, KSWork.KS_MEASURE_COL, this.RegisterSheet);

            ks_work.Number = number;
            var estimate_namber = registerSheet.Cells[rowIndex, KSWork.KS_ESTIMATE_NUMBER_COL].Value;
            if (estimate_namber != null)
                ks_work.EstimateNumber = estimate_namber.ToString();


            var code = registerSheet.Cells[rowIndex, KSWork.KS_CODE_COL].Value;
            if (code != null)
                ks_work.Code = code.ToString();
            else
                ks_work.SetPropertyValidStatus("Code", false);

            var name = registerSheet.Cells[rowIndex, KSWork.KS_NAME_COL].Value;
            if (name != null)
                ks_work.Name = name.ToString();
            else
                ks_work.SetPropertyValidStatus("Name", false);

            var unit_of_measurement_name = registerSheet.Cells[rowIndex, KSWork.KS_MEASURE_COL].Value;
            if (unit_of_measurement_name != null)
                ks_work.UnitOfMeasurement = new UnitOfMeasurement(unit_of_measurement_name);
            else
                ks_work.SetPropertyValidStatus("UnitOfMeasurement.Name", false);

            var pr_quantity = registerSheet.Cells[rowIndex, KSWork.KS_QUANTITY_COL].Value;
            if (pr_quantity != null && pr_quantity != 0)
                ks_work.ProjectQuantity = Decimal.Parse(pr_quantity.ToString());
            else
                ks_work.SetPropertyValidStatus("ProjectQuantity", false);

            var quantity = registerSheet.Cells[rowIndex, KSWork.KS_QUANTITY_FACT_COL].Value;
            if (quantity != null)
                ks_work.Quantity = Decimal.Parse(quantity.ToString());

            var laboriousness = registerSheet.Cells[rowIndex, KSWork.KS_LABOURNESS_COL].Value;
            if (laboriousness != null && laboriousness != 0)
                ks_work.Laboriousness = Decimal.Parse(laboriousness.ToString());
            else
                ks_work.SetPropertyValidStatus("Laboriousness", false);


            if (!this.KSWorks.Contains(ks_work))
                this.KSWorks.Add(ks_work);
        }
        /// <summary>
        /// Функция из части КС-2 листа Worksheet создает и помещает в модель работы типа KSWork 
        /// </summary>
        public void LoadKSWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;

            null_str_count = 0;
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, KSWork.KS_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    this.LoadKSWork(rowIndex);
                }
                rowIndex++;
            }
        }


        public void LoadRCWork(int row)
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = row;

            var number = registerSheet.Cells[rowIndex, RCWork.RC_NUMBER_COL].Value;
            if (number == null) return;

            RCWork rc_work = this.RCWorks.FirstOrDefault(w => w.Number == number);
            if (rc_work == null)
                rc_work = new RCWork();

            rc_work.LoadSessionId = this.LoadSessionId;
            rc_work.Worksheet = registerSheet;
            this.Register(rc_work, "Number", rowIndex, RCWork.RC_NUMBER_COL, this.RegisterSheet, false, v => Regex.IsMatch(v.ToString(), @"^\d+\.\d+\.\d+\.\d+\.\d+$"));
            this.Register(rc_work, "Code", rowIndex, RCWork.RC_CODE_COL, this.RegisterSheet);
            this.Register(rc_work, "Name", rowIndex, RCWork.RC_NAME_COL, this.RegisterSheet);
            this.Register(rc_work, "ProjectQuantity", rowIndex, RCWork.RC_QUANTITY_COL, this.RegisterSheet);
            this.Register(rc_work, "Quantity", rowIndex, RCWork.RC_QUANTITY_FACT_COL, this.RegisterSheet, true);
            this.Register(rc_work, "LabournessCoefficient", rowIndex, RCWork.RC_LABOURNESS_COEFFICIENT_COL, this.RegisterSheet);
            this.Register(rc_work, "Laboriousness", rowIndex, RCWork.RC_LABOURNESS_COL, this.RegisterSheet);
            this.Register(rc_work, "UnitOfMeasurement.Name", rowIndex, RCWork.RC_MEASURE_COL, this.RegisterSheet);


            rc_work.Number = number;
            var code = registerSheet.Cells[rowIndex, RCWork.RC_CODE_COL].Value;
            if (code != null)
                rc_work.Code = code;
            else
                rc_work.SetPropertyValidStatus("Code", false);

            var name = registerSheet.Cells[rowIndex, RCWork.RC_NAME_COL].Value;
            if (name != null)
                rc_work.Name = name;
            else
                rc_work.SetPropertyValidStatus("Name", false);

            var unit_of_measurement_name = registerSheet.Cells[rowIndex, RCWork.RC_MEASURE_COL].Value;
            if (unit_of_measurement_name != null)
                rc_work.UnitOfMeasurement = new UnitOfMeasurement(unit_of_measurement_name);
            else
                rc_work.SetPropertyValidStatus("UnitOfMeasurement.Name", false);

            var pr_quantity = registerSheet.Cells[rowIndex, RCWork.RC_QUANTITY_COL].Value;
            if (pr_quantity != null && pr_quantity != 0)
                rc_work.ProjectQuantity = Decimal.Parse(pr_quantity.ToString());
            else
                rc_work.SetPropertyValidStatus("ProjectQuantity", false);

            var quantity = registerSheet.Cells[rowIndex, RCWork.RC_QUANTITY_FACT_COL].Value;
            if (quantity != null)
                rc_work.Quantity = Decimal.Parse(quantity.ToString());

            var laboriosness_coef = registerSheet.Cells[rowIndex, RCWork.RC_LABOURNESS_COEFFICIENT_COL].Value;
            if (laboriosness_coef != null)
                rc_work.LabournessCoefficient = Decimal.Parse(laboriosness_coef.ToString());
            else
            {
                rc_work.LabournessCoefficient = 0;
                //  rc_work.SetPropertyValidStatus("LabournessCoefficient", false);

            }
            var laboriousness = registerSheet.Cells[rowIndex, RCWork.RC_LABOURNESS_COL].Value;
            if (laboriousness != null && laboriousness != 0)
                rc_work.Laboriousness = Decimal.Parse(laboriousness.ToString());
            else
            {
                rc_work.Laboriousness = 0;
                //    rc_work.SetPropertyValidStatus("Laboriousness", false);

            }

            if (!this.RCWorks.Contains(rc_work))
                this.RCWorks.Add(rc_work);

        }
        /// <summary>
        /// Функция из части КС-2 листа Worksheet создает и помещает в модель работы типа KSWork 
        /// </summary>
        public void LoadRCWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;


            null_str_count = 0;
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, RCWork.RC_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;

                    this.LoadRCWork(rowIndex);
                }
                rowIndex++;
            }
        }

        public void LoadWorksReport(int row)
        {
            int rowIndex = row;
            Excel.Worksheet registerSheet = this.RegisterSheet;
            string rc_number = registerSheet.Cells[rowIndex, WorkReportCard.WRC_NUMBER_COL].Value.ToString(); ;

            WorkReportCard report_card = this.WorkReportCards.FirstOrDefault(rc => rc.Number == rc_number);
            if (report_card == null)
                report_card = new WorkReportCard();
            else
                report_card.Clear();

            report_card.LoadSessionId = this.LoadSessionId;
            report_card.Worksheet = registerSheet;
            DateTime end_date = this.WorksEndDate;
            report_card.Number = rc_number;
            this.Register(report_card, "Number", rowIndex, WorkReportCard.WRC_NUMBER_COL, this.RegisterSheet);
            this.Register(report_card, "PreviousComplatedQuantity", rowIndex, WorkReportCard.WRC_PC_QUANTITY_COL, this.RegisterSheet, true);

            var previus_comp_quantity = registerSheet.Cells[rowIndex, WorkReportCard.WRC_PC_QUANTITY_COL].Value;
            if (previus_comp_quantity != null)
                report_card.PreviousComplatedQuantity = Decimal.Parse(previus_comp_quantity.ToString());

            int date_index = 0;
            while (date_index < this.WorkedDaysNumber)
            {
                DateTime current_date = DateTime.Parse(registerSheet.Cells[WorkReportCard.WRC_DATE_ROW, WorkReportCard.WRC_DATE_COL + date_index].Value.ToString());

                decimal quantity = 0;
                if (registerSheet.Cells[rowIndex, WorkReportCard.WRC_DATE_COL + date_index].Value != null)
                    quantity = Decimal.Parse(registerSheet.Cells[rowIndex, WorkReportCard.WRC_DATE_COL + date_index].Value.ToString());
                if (quantity != 0)
                {
                    WorkDay workDay = new WorkDay();
                    workDay.Worksheet = registerSheet;
                    workDay.Date = current_date;
                    workDay.Quantity = quantity;
                    this.Register(workDay, "Quantity", rowIndex, WorkReportCard.WRC_DATE_COL + date_index, this.RegisterSheet);
                    report_card.Add(workDay);
                }
                date_index++;
            }
            report_card.DaysFirsDate = this.RecordCardStartDate;
            if (!this.WorkReportCards.Contains(report_card))
                this.WorkReportCards.Add(report_card);
        }
        /// <summary>
        /// Функция из календарной части (левой части) листа Worksheet создает и помещает в соответсвующие  
        /// работы типа KSWork табели выполненных работ ReportCard с объектами типа WorkDay с даной и количеством 
        /// выполенной работы
        /// </summary>
        public void LoadWorksReportCards()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, WorkReportCard.WRC_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    this.LoadWorksReport(rowIndex);
                }
                rowIndex++;
            }

        }
        public void LoadWorkerConsumptions()
        {
            Excel.Worksheet consumtionsSheet = this.WorkerConsumptionsSheet;
            int rowIndex = WorkerConsumption.W_CONSUMPTIONS_FIRST_ROW_INDEX;
            foreach (var wc in this.WorkerConsumptions)
                this.Unregister(wc);
            this.WorkerConsumptions.Clear();
            null_str_count = 0;

            while (null_str_count < 100)
            {
                var number = consumtionsSheet.Cells[rowIndex, WorkerConsumption.W_CONSUMPTIONS_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    WorkerConsumption worker_consumption = new WorkerConsumption();
                    worker_consumption.Worksheet = consumtionsSheet;
                    this.Register(worker_consumption, "Number", rowIndex, WorkerConsumption.W_CONSUMPTIONS_NUMBER_COL, consumtionsSheet);
                    this.Register(worker_consumption, "Name", rowIndex, WorkerConsumption.W_CONSUMPTIONS_NAME_COL, consumtionsSheet);

                    worker_consumption.Number = number.ToString();
                    worker_consumption.Name = consumtionsSheet.Cells[rowIndex, WorkerConsumption.W_CONSUMPTIONS_NAME_COL].Value; ;
                    worker_consumption.WorkersConsumptionReportCard.Clear();

                    int date_index = 0;

                    while (date_index < this.WorkedDaysNumber)
                    {
                        DateTime current_date = DateTime.Parse(consumtionsSheet.Cells[WorkerConsumption.W_CONSUMPTIONS_DATE_RAW, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString());
                        decimal quantity = 0;
                        if (consumtionsSheet.Cells[rowIndex, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value != null)
                            quantity = Decimal.Parse(consumtionsSheet.Cells[rowIndex, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString());

                        if (quantity != 0)
                        {
                            WorkerConsumptionDay w_consumption_Day = new WorkerConsumptionDay();
                            w_consumption_Day.Worksheet = consumtionsSheet;
                            w_consumption_Day.Date = current_date;
                            w_consumption_Day.Quantity = quantity;
                            this.Register(w_consumption_Day, "Quantity", rowIndex, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + date_index, consumtionsSheet);
                            worker_consumption.WorkersConsumptionReportCard.Add(w_consumption_Day);
                        }

                        date_index++;
                    }
                    worker_consumption.WorkersConsumptionReportCard.DaysFirsDate = this.RecordCardStartDate;
                    if (!this.WorkerConsumptions.Contains(worker_consumption))
                        this.WorkerConsumptions.Add(worker_consumption);

                }
                rowIndex++;
            }

        }
        public void LoadMachineConsumptions()
        {
            Excel.Worksheet consumtionsSheet = this.MachineConsumptionsSheet;
            int rowIndex = MachineConsumption.MCH_CONSUMPTIONS_FIRST_ROW_INDEX;
            foreach (var mc in this.MachineConsumptions)
                this.Unregister(mc);
            this.MachineConsumptions.Clear();
            null_str_count = 0;

            while (null_str_count < 100)
            {
                var number = consumtionsSheet.Cells[rowIndex, MachineConsumption.MCH_CONSUMPTIONS_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    MachineConsumption machine_consumption = new MachineConsumption();
                    machine_consumption.Worksheet = consumtionsSheet;
                    this.Register(machine_consumption, "Number", rowIndex, MachineConsumption.MCH_CONSUMPTIONS_NUMBER_COL, consumtionsSheet);
                    this.Register(machine_consumption, "Name", rowIndex, MachineConsumption.MCH_CONSUMPTIONS_NAME_COL, consumtionsSheet);

                    machine_consumption.Number = number.ToString();
                    var name = consumtionsSheet.Cells[rowIndex, MachineConsumption.MCH_CONSUMPTIONS_NAME_COL].Value;
                    machine_consumption.Name = name;
                    machine_consumption.MachinesConsumptionReportCard.Clear();

                    int date_index = 0;

                    while (date_index < this.WorkedDaysNumber)
                    {
                        DateTime current_date = DateTime.Parse(consumtionsSheet.Cells[MachineConsumption.MCH_CONSUMPTIONS_DATE_RAW, MachineConsumption.MCH_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString());
                        decimal quantity = 0;
                        if (consumtionsSheet.Cells[rowIndex, MachineConsumption.MCH_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value != null)
                            quantity = Decimal.Parse(consumtionsSheet.Cells[rowIndex, MachineConsumption.MCH_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString());

                        if (quantity != 0)
                        {
                            MachineConsumptionDay w_consumption_Day = new MachineConsumptionDay();
                            w_consumption_Day.Date = current_date;
                            w_consumption_Day.Quantity = quantity;
                            this.Register(w_consumption_Day, "Quantity", rowIndex, MachineConsumption.MCH_CONSUMPTIONS_FIRST_DATE_COL + date_index, consumtionsSheet);
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
        public void ClearAllSections()
        {
            this.ClearAllMSGWorks();
            foreach (var section in this.WorksSections)
                this.Unregister(section);
            this.WorksSections.Clear();
        }
        public void ClearAllMSGWorks()
        {
            this.ClearAllVOVRWorks();
            foreach (var work in this.MSGWorks)
            {
                foreach (WorkScheduleChunk ch in work.WorkSchedules)
                    this.Unregister(ch);
                foreach (NeedsOfWorker nw in work.WorkersComposition)
                    this.Unregister(nw);
                foreach (NeedsOfMachine nm in work.MachinesComposition)
                    this.Unregister(nm);
                work.WorkSchedules.Clear();
                work.WorkersComposition.Clear();
                work.MachinesComposition.Clear();
                this.Unregister(work);
            }

            this.MSGWorks.Clear();
            this.WorkersComposition.Clear();
            this.MachinesComposition.Clear();
        }
        public void ClearAllVOVRWorks()
        {
            this.ClearAllKSWorks();
            foreach (var work in this.VOVRWorks)
                this.Unregister(work);
            this.VOVRWorks.Clear();
            foreach (var w in this.MSGWorks)
                w.VOVRWorks.Clear();

        }
        public void ClearAllKSWorks()
        {
            this.ClearAllRCWorks();
            foreach (var work in this.KSWorks)
                this.Unregister(work);
            this.KSWorks.Clear();
            foreach (var w in this.VOVRWorks)
                w.KSWorks.Clear();
        }
        public void ClearAllRCWorks()
        {
            foreach (var work in this.RCWorks)
                this.Unregister(work);
            this.RCWorks.Clear();
            foreach (var w in this.KSWorks)
                w.RCWorks.Clear();
        }
        public void ClearAllRecordCards()
        {
            List<WorkReportCard> all_rcards = new List<WorkReportCard>(this.WorkReportCards);
            foreach (var rc in all_rcards)
            {
                this.WorkReportCards.Remove(rc);
                this.Unregister(rc);
            }
            this.WorkReportCards.Clear();
        }

        public void ClearAllWorkerConsuptions()
        {
            List<WorkerConsumption> all_wconsumpts = new List<WorkerConsumption>(this.WorkerConsumptions);
            foreach (var wc in all_wconsumpts)
            {
                this.WorkerConsumptions.Remove(wc);
                this.Unregister(wc);
            }
            this.WorkerConsumptions.Clear();
        }
        public void ClearAllMachineConsuptions()
        {
            List<MachineConsumption> all_mconsumpts = new List<MachineConsumption>(this.MachineConsumptions);
            foreach (var mc in all_mconsumpts)
            {
                this.MachineConsumptions.Remove(mc);
                this.Unregister(mc);
            }
            this.MachineConsumptions.Clear();
        }

        public void AdjustObjectModel()
        {
            foreach (WorksSection section in this.WorksSections.ToList())
            {
                if (section.LoadSessionId != this.LoadSessionId)
                {
                    this.Unregister(section);
                    if (section.Owner != null && section.Owner is MSGExellModel model)
                        model.WorksSections.Remove(section);

                    this.WorksSections.Remove(section);
                }
                var sections = this.WorksSections.Where(s => s.Number == section.Number && !InvalidObjects.Contains(s)).ToList();
                if (sections.Count > 1)
                {
                    sections[0].SetPropertyValidStatus("Number", false);
                    sections.Remove(sections[0]);
                    foreach (WorksSection s in sections)
                    {
                        s.SetPropertyValidStatus("Number", false);
                        InvalidObjects.Add(s);
                        this.WorksSections.Remove(s);
                    }
                }
            }
            foreach (MSGWork work in this.MSGWorks.ToList())
            {
                if (work.LoadSessionId != this.LoadSessionId)
                {
                    this.Unregister(work);
                    if (work.Owner != null && work.Owner is WorksSection section)
                        section.MSGWorks.Remove(work);
                    this.MSGWorks.Remove(work);
                }

                var works = this.MSGWorks.Where(w => w.Number == work.Number && !InvalidObjects.Contains(w)).ToList();
                if (works.Count > 1)
                {
                    works[0].SetPropertyValidStatus("Number", false);
                    works.Remove(works[0]);
                    foreach (MSGWork w in works)
                    {
                        w.SetPropertyValidStatus("Number", false);
                        InvalidObjects.Add(w);
                        this.MSGWorks.Remove(w);
                    }
                }
            }
            foreach (VOVRWork work in this.VOVRWorks.ToList())
            {
                if (work.LoadSessionId != this.LoadSessionId)
                {
                    this.Unregister(work);
                    if (work.Owner != null && work.Owner is MSGWork msg_work)
                        msg_work.VOVRWorks.Remove(work);
                    this.VOVRWorks.Remove(work);
                }

                var works = this.VOVRWorks.Where(w => w.Number == work.Number && !InvalidObjects.Contains(w)).ToList();
                if (works.Count > 1)
                {
                    works[0].SetPropertyValidStatus("Number", false);
                    works.Remove(works[0]);
                    foreach (VOVRWork w in works)
                    {
                        w.SetPropertyValidStatus("Number", false);
                        InvalidObjects.Add(w);
                        this.VOVRWorks.Remove(w);
                    }
                }
            }
            foreach (KSWork work in this.KSWorks.ToList())
            {
                if (work.LoadSessionId != this.LoadSessionId)
                {
                    this.Unregister(work);
                    if (work.Owner != null && work.Owner is VOVRWork vovr_work)
                        vovr_work.KSWorks.Remove(work);
                    this.KSWorks.Remove(work);
                }
                var works = this.KSWorks.Where(w => w.Number == work.Number && !InvalidObjects.Contains(w)).ToList();
                if (works.Count > 1)
                {
                    works[0].SetPropertyValidStatus("Number", false);
                    works.Remove(works[0]);
                    foreach (KSWork w in works)
                    {
                        w.SetPropertyValidStatus("Number", false);
                        InvalidObjects.Add(w);
                        this.KSWorks.Remove(w);
                    }
                }
            }
            foreach (RCWork work in this.RCWorks.ToList())
            {
                if (work.LoadSessionId != this.LoadSessionId)
                {
                    this.Unregister(work);
                    if (work.Owner != null && work.Owner is KSWork ks_work)
                        ks_work.RCWorks.Remove(work);

                    this.RCWorks.Remove(work);
                }
                var works = this.RCWorks.Where(w => w.Number == work.Number && !InvalidObjects.Contains(w)).ToList();
                if (works.Count > 1)
                {
                    works[0].SetPropertyValidStatus("Number", false);
                    works.Remove(works[0]);
                    foreach (RCWork w in works)
                    {
                        w.SetPropertyValidStatus("Number", false);
                        InvalidObjects.Add(w);
                        this.RCWorks.Remove(w);
                    }
                }
            }
            foreach (WorkReportCard rcard in this.WorkReportCards.ToList())
            {
                if (rcard.LoadSessionId != this.LoadSessionId)
                {
                    this.Unregister(rcard);
                    this.WorkReportCards.Remove(rcard);
                }

                var rcards = this.WorkReportCards.Where(w => w.Number == rcard.Number && !InvalidObjects.Contains(w)).ToList();
                if (rcards.Count > 1)
                {
                    rcards[0].SetPropertyValidStatus("Number", false);
                    rcards.Remove(rcards[0]);
                    foreach (WorkReportCard rc in rcards)
                    {
                        rc.SetPropertyValidStatus("Number", false);
                        InvalidObjects.Add(rc);
                        this.WorkReportCards.Remove(rc);
                    }
                }

            }

            foreach (MSGWork msg_work in this.MSGWorks.OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberPrefix}.", ""))))
            {
                WorksSection w_section = this.WorksSections.Where(ws => ws.Number == msg_work.NumberPrefix).FirstOrDefault();
                if (w_section != null)
                {
                    w_section.MSGWorks.Owner = w_section;
                    //  msg_work.Owner = w_section;
                    if (!w_section.MSGWorks.Contains(msg_work))
                        w_section.MSGWorks.Add(msg_work);

                }
                foreach (VOVRWork vovr_work in this.VOVRWorks.Where(w => w.NumberPrefix == msg_work.Number).OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberPrefix}.", ""))))
                {
                    VOVRWork finded_vovr_work = msg_work.VOVRWorks.FirstOrDefault(vr_w => vr_w.Number == vovr_work.Number);
                    if (finded_vovr_work == null)
                    {
                        msg_work.VOVRWorks.Owner = msg_work;
                        //     vovr_work.Owner = msg_work;
                        if (!msg_work.VOVRWorks.Contains(vovr_work))
                            msg_work.VOVRWorks.Add(vovr_work);

                    }

                    foreach (KSWork ks_work in this.KSWorks.Where(w => w.NumberPrefix == vovr_work.Number).OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberPrefix}.", ""))))
                    {
                        KSWork finded_ks_work = vovr_work.KSWorks.FirstOrDefault(kcw => kcw.Number == ks_work.Number);
                        if (finded_ks_work == null)
                        {
                            vovr_work.KSWorks.Owner = vovr_work;
                            //ks_work.Owner = vovr_work;
                            if (!vovr_work.KSWorks.Contains(ks_work))
                                vovr_work.KSWorks.Add(ks_work);

                        }

                        foreach (RCWork rc_work in this.RCWorks.Where(w => w.NumberPrefix == ks_work.Number).OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberPrefix}.", ""))))
                        {
                            RCWork finded_rc_work = ks_work.RCWorks.FirstOrDefault(rcw => rcw.Number == rc_work.Number);
                            if (finded_rc_work == null)
                            {
                                ks_work.RCWorks.Owner = ks_work;
                                //   rc_work.Owner = ks_work;
                                if (!ks_work.RCWorks.Contains(rc_work))
                                    ks_work.RCWorks.Add(rc_work);

                            }

                            var report_card = this.WorkReportCards.Where(r => r.Number == rc_work.Number).FirstOrDefault();
                            if (report_card != null)
                            {
                                report_card.Owner = rc_work;
                                report_card.Worksheet = rc_work.Worksheet;
                                rc_work.ReportCard = report_card;
                            }

                        }


                    }
                }
            }
        }
        private void AdjustRCWorksRecorCard()
        {
            foreach (RCWork rc_work in this.RCWorks)
            {
                WorkReportCard finded_rc = this.WorkReportCards.FirstOrDefault(rc => rc.Number == rc_work.Number);
                if (finded_rc != null)
                {
                    finded_rc.Owner = rc_work;
                    finded_rc.Worksheet = rc_work.Worksheet;
                    rc_work.ReportCard = finded_rc;

                }
            }
        }
        /// <summary>
        /// Заргужает(перезагружает)  данныхе из соотвествующих листов в модель
        /// </summary>
        public void ReloadSheetModel()
        {

            //  this.WorksStartDate = DateTime.Parse(this.RegisterSheet.Cells[WORKS_START_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
            //   this.Register(this, "WorksStartDate", WORKS_START_DATE_ROW, WORKS_END_DATE_COL, this.RegisterSheet);
            this.RecordCardStartDate = DateTime.Parse(this.RegisterSheet.Cells[WORKS_START_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
            this.Register(this, "RecordCardStartDate", WORKS_START_DATE_ROW, WORKS_END_DATE_COL, this.RegisterSheet);


            this.ContractCode = this.CommonSheet.Cells[CONTRACT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
            this.ContructionObjectCode = this.CommonSheet.Cells[CONSTRUCTION_OBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
            this.ConstructionSubObjectCode = this.CommonSheet.Cells[CONSTRUCTION_SUBOBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
            //this.CellAddressesMap.Add("ContractCode", new ExellPropAddress<int, int, Worksheet>(CONTRACT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));
            //this.CellAddressesMap.Add("ContructionObjectCode", new ExellPropAddress<int, int, Worksheet>(CONSTRUCTION_OBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));
            //this.CellAddressesMap.Add("ConstructionSubObjectCode", new ExellPropAddress<int, int, Worksheet>(CONSTRUCTION_SUBOBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));
            this.AllHashDictationary.Clear();


            this.WorksStartDate = DateTime.Parse(this.RegisterSheet.Cells[WORKS_START_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
            this.InvalidObjects.Clear();

            if (IsHasEnabled)
                this.SetHashFormulas();
            //this.ClearAllSections();
            this.LoadSessionId = Guid.NewGuid();

            this.LoadWorksReportCards();

            this.LoadWorksSections();
            this.LoadMSGWorks();
            this.LoadVOVRWorks();
            this.LoadKSWorks();
            this.LoadRCWorks();
            this.AdjustObjectModel();

            this.LoadWorkerConsumptions();
            this.LoadMachineConsumptions();
            this.CalcLabourness();
            this.WorksSections.Validate();
        }
        public void ReadModelFilds()
        {
            foreach (WorksSection section in this.WorksSections)
            {
                section.LoadExellBindableObjectFromField();
                foreach (var msg_work in section.MSGWorks)
                {
                    msg_work.LoadExellBindableObjectFromField();
                    foreach (var vovr_work in msg_work.VOVRWorks)
                    {
                        vovr_work.LoadExellBindableObjectFromField();
                        foreach (var ks_work in vovr_work.KSWorks)
                        {
                            ks_work.LoadExellBindableObjectFromField();
                            foreach (var rc_work in ks_work.RCWorks)
                            {
                                rc_work.LoadExellBindableObjectFromField();
                            }
                        }
                    }
                }

            }
        }

        public void SetHashFormulas()
        {
            if (this.Owner != null) return;
            int null_str_count = 0;
            int rowIndex = FIRST_ROW_INDEX;
            int columnIndex = WorksSection.WSEC_NUMBER_COL;
            Excel.Worksheet registerSheet = this.RegisterSheet;
            RowsHashValues.Clear();
            ColumnsHashValues.Clear();
            while (null_str_count < 100)
            {
                var number = registerSheet.Cells[rowIndex, WorkReportCard.WRC_NUMBER_COL].Value;
                Excel.Range row_hash_range = registerSheet.Cells[rowIndex, HASH_FUNCTION_COL];
                Excel.Range first_cell = registerSheet.Cells[rowIndex, WorksSection.WSEC_NUMBER_COL];
                Excel.Range last_cell = registerSheet.Cells[rowIndex, RCWork.RC_LABOURNESS_COL];
                row_hash_range.NumberFormat = "0";
                row_hash_range.Formula = $"= EasyHash(CONCAT({first_cell.RangeAddress()}:{last_cell.RangeAddress()}))";//= EasyHash(СЦЕП(B8: AK8))
                var row_hash_val = registerSheet.Cells[rowIndex, HASH_FUNCTION_COL].Value;
                if (row_hash_val != null)
                {
                    int row = Int32.Parse(row_hash_val.ToString());
                    //  if (!this.RowsHashValues.Contains(row))
                    this.RowsHashValues.Add(row);
                }

                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;

                }
                rowIndex++;
            }

            for (int colIndex = WorksSection.WSEC_NUMBER_COL; colIndex <= RCWork.RC_LABOURNESS_COL; colIndex++)
            {
                Excel.Range col_hash_range = registerSheet.Cells[HASH_FUNCTION_ROW, colIndex];
                Excel.Range first_cell = registerSheet.Cells[FIRST_ROW_INDEX, colIndex];
                Excel.Range last_cell = registerSheet.Cells[FIRST_ROW_INDEX + MAX_HASH_FUNCTION_ROW, colIndex];
                col_hash_range.Formula = $"= EasyHash(CONCAT({first_cell.RangeAddress()}:{last_cell.RangeAddress()}))";
                col_hash_range.NumberFormat = "0";
                var col_hash_val = registerSheet.Cells[HASH_FUNCTION_ROW, colIndex].Value;
                if (col_hash_val != null)
                {
                    int col = Int32.Parse(col_hash_val.ToString());
                    //  if (!ColumnsHashValues.Contains(col))
                    ColumnsHashValues.Add(col);
                }
            }

        }


        public ObservableCollection<IObservableExcelBindableBase> ChangedObjectsCollection = new ObservableCollection<IObservableExcelBindableBase>();

        /// <summary>
        /// Функция находит строки и стобцы с измененными данными и обновляет объекты которые попадают в эти строки и стобцы
        /// </summary>
        public void ReloadSheetModelLocal()
        {
            int null_str_count = 0;
            int rowIndex = FIRST_ROW_INDEX;
            int columnIndex = WorksSection.WSEC_NUMBER_COL;
            Excel.Worksheet registerSheet = this.RegisterSheet;
            ObservableCollection<Tuple<int, int>> changed_objects_corredinates = new ObservableCollection<Tuple<int, int>>();
            List<int> current_hash_columns = new List<int>();

            for (int colIndex = WorksSection.WSEC_NUMBER_COL; colIndex <= RCWork.RC_LABOURNESS_COL; colIndex++)
            {
                var col_hash_val = registerSheet.Cells[HASH_FUNCTION_ROW, colIndex].Value;
                if (col_hash_val != null)
                {
                    int col_hash = Int32.Parse(col_hash_val.ToString());
                    current_hash_columns.Add(col_hash);
                }
            }
            List<int> _columns_with_changes = new List<int>();
            foreach (int col_hash in current_hash_columns)
            {

                int col_index = current_hash_columns.IndexOf(col_hash);
                int col_indexs = col_index + WorksSection.WSEC_NUMBER_COL;
                if (!ColumnsHashValues.Contains(col_hash))
                    _columns_with_changes.Add(col_indexs);
            }

            List<int> _rows_with_changes = new List<int>();
            while (null_str_count < 100)
            {
                var row_hash_val = registerSheet.Cells[rowIndex, HASH_FUNCTION_COL].Value;
                if (row_hash_val != null)
                {
                    int row_hash = Int32.Parse(row_hash_val.ToString());
                    if (!RowsHashValues.Contains(row_hash))
                    {
                        int row_index = RowsHashValues.IndexOf(row_hash);
                        int row = rowIndex;
                        _rows_with_changes.Add(row);
                        //foreach (int col_hash in current_hash_columns)
                        //{

                        //    int col_index = current_hash_columns.IndexOf(col_hash);
                        //    int col_indexs = col_index + WorksSection.WSEC_NUMBER_COL;
                        //    if (!ColumnsHashValues.Contains(col_hash))
                        //        changed_objects_corredinates.Add(new Tuple<int, int>(row, col_indexs));
                        //}
                    }
                }
                var number = registerSheet.Cells[rowIndex, WorkReportCard.WRC_NUMBER_COL].Value;
                if (number == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                }
                rowIndex++;
            }
            int top_row = _rows_with_changes.OrderBy(r => r).FirstOrDefault();
            int bottom_row = _rows_with_changes.OrderBy(r => r).LastOrDefault();

            if (_columns_with_changes.Where(c => c >= WorksSection.WSEC_NUMBER_COL && c <= WorksSection.WSEC_NAME_COL).Any())
            {
                if (top_row == 0 || bottom_row == 0)
                    this.LoadWorksSections();
                else
                    for (int row = top_row; row <= bottom_row; row++)
                        this.LoadWorksSection(row);
            }

            if (_columns_with_changes.Where(c => c >= MSGWork.MSG_NUMBER_COL && c <= MSGWork.MSG_NEEDS_OF_MACHINE_QUANTITY_COL).Any())
            {
                if (top_row == 0 || bottom_row == 0)
                    this.LoadMSGWorks();
                else
                    for (int row = top_row; row <= bottom_row; row++)
                        this.LoadMSGWork(row);
            }

            if (_columns_with_changes.Where(c => c >= VOVRWork.VOVR_NUMBER_COL && c <= VOVRWork.VOVR_LABOURNESS_COL).Any())
            {
                if (top_row == 0 || bottom_row == 0)
                    this.LoadKSWorks();
                else
                    for (int row = top_row; row <= bottom_row; row++)
                        this.LoadVOVRWork(row);
            }

            if (_columns_with_changes.Where(c => c >= KSWork.KS_NUMBER_COL && c <= KSWork.KS_LABOURNESS_COL).Any())
            {
                if (top_row == 0 || bottom_row == 0)
                    this.LoadKSWorks();
                else
                    for (int row = top_row; row <= bottom_row; row++)
                        this.LoadKSWork(row);
            }

            if (_columns_with_changes.Where(c => c >= RCWork.RC_NUMBER_COL && c <= RCWork.RC_LABOURNESS_COL).Any())
            {
                if (top_row == 0 || bottom_row == 0)
                    this.LoadRCWorks();
                else
                    for (int row = top_row; row <= bottom_row; row++)
                        this.LoadRCWork(row);
            }


            this.SetHashFormulas();
            this.AdjustObjectModel();
        }


        //private void ReloadLoadExbbObject(IObservableExcelBindableBase obj)
        //{
        //    foreach (var kvp in obj.CellAddressesMap)
        //    {
        //        var prop_infoes = obj.GetType().GetProperties().Where(p => p.Name == kvp.Key);
        //        foreach (PropertyInfo prop_info in prop_infoes)
        //        {
        //            var prop_val = prop_info.GetValue(obj);
        //            prop_info.SetValue(obj, kvp.Value.Cell.Value);
        //        }
        //    }
        //}

        /// <summary>
        /// Функция форматирует представления модели на листе Excel
        /// </summary>
        public override void SetStyleFormats(int col = W_SECTION_COLOR)
        {
            //  this.UpdateCellAddressMapsWorkSheets();
            this.RemoveGroups(this.RegisterSheet);
            int selectin_col = col;
            if (this.WorksSections.Count > 0)
            {
                //  Excel.Range _sections_left_edge_range = this.WorksSections.Worksheet.Range[this.WorksSections[0].CellAddressesMap["Number"].Cell,
                //                                                            this.WorksSections[this.WorksSections.Count - 1].CellAddressesMap["Number"].Cell];
                //          _sections_left_edge_range.SetBordersLine(XlLineStyle.xlLineStyleNone, XlLineStyle.xlDashDot, XlLineStyle.xlLineStyleNone, XlLineStyle.xlLineStyleNone);

            }

            foreach (WorksSection section in this.WorksSections)
                section.SetStyleFormats(selectin_col);

            foreach (IObservableExcelBindableBase obj in InvalidObjects)
                obj.SetStyleFormats(selectin_col);

            this.WorksSections.SetInvalidateCellsColor(XlRgbColor.rgbRed);
            this.MSGWorks.SetInvalidateCellsColor(XlRgbColor.rgbRed);
            this.VOVRWorks.SetInvalidateCellsColor(XlRgbColor.rgbRed);
            this.KSWorks.SetInvalidateCellsColor(XlRgbColor.rgbRed);
            this.RCWorks.SetInvalidateCellsColor(XlRgbColor.rgbRed);
            this.WorkReportCards.SetInvalidateCellsColor(XlRgbColor.rgbRed);
            this.InvalidObjects.SetInvalidateCellsColor(XlRgbColor.rgbRed);

            this.WorkerConsumptions.GetRange().SetBordersLine();
            int w_consumption_col = W_SECTION_COLOR;
            foreach (WorkerConsumption consumption in this.WorkerConsumptions)
            {
                // consumption.GetRange(this.WorkerConsumptionsSheet).Interior.ColorIndex = w_consumption_col++;
                int days_namber = (this.WorksEndDate - this.RecordCardStartDate).Days;
                Excel.Range cons_range = this.WorkerConsumptionsSheet.Range[
                    this.WorkerConsumptionsSheet.Cells[consumption.GetTopRow(), WorkerConsumption.W_CONSUMPTIONS_NUMBER_COL],
                    this.WorkerConsumptionsSheet.Cells[consumption.GetTopRow(), days_namber]];
                cons_range.Interior.ColorIndex = w_consumption_col++;
                cons_range.Borders.LineStyle = XlLineStyle.xlDashDotDot;
                cons_range.SetBordersLine(XlLineStyle.xlDouble, XlLineStyle.xlDouble,
                    XlLineStyle.xlContinuous, XlLineStyle.xlContinuous);
            }
            w_consumption_col = W_SECTION_COLOR;
            foreach (MachineConsumption consumption in this.MachineConsumptions)
            {
                // consumption.GetRange(this.WorkerConsumptionsSheet).Interior.ColorIndex = w_consumption_col++;
                int days_namber = (this.WorksEndDate - this.RecordCardStartDate).Days;
                Excel.Range cons_range = this.MachineConsumptionsSheet.Range[
                    this.MachineConsumptionsSheet.Cells[consumption.GetTopRow(), MachineConsumption.MCH_CONSUMPTIONS_NUMBER_COL],
                    this.MachineConsumptionsSheet.Cells[consumption.GetTopRow(), days_namber]];
                cons_range.Interior.ColorIndex = w_consumption_col++;
                cons_range.Borders.LineStyle = XlLineStyle.xlDashDotDot;
                cons_range.SetBordersLine(XlLineStyle.xlDouble, XlLineStyle.xlDouble,
                    XlLineStyle.xlContinuous, XlLineStyle.xlContinuous);
            }

            Excel.Range vovr_colomns = this.RegisterSheet.Range[this.RegisterSheet.Columns[VOVRWork.VOVR_NUMBER_COL], this.RegisterSheet.Columns[VOVRWork.VOVR_LABOURNESS_COL]];
            Excel.Range ks_colomns = this.RegisterSheet.Range[this.RegisterSheet.Columns[VOVRWork.VOVR_NUMBER_COL], this.RegisterSheet.Columns[KSWork.KS_LABOURNESS_COL]];
            try
            {
                ks_colomns.Group();
                vovr_colomns.Group();
            }
            catch (Exception exp)
            {
                throw new Exception($"Ошибка при группировке стобцов документа. Метод MSGExcelModel.SetStyleFormats(..). {this.ToString()}: {this.Number}.Ошибка:{exp.Message}");
            }

        }

        /// <summary>
        ///Фунция проставляет все соотвесвующие формулы в ячейках Excell в соотвествии с моделью
        /// </summary>
        public void SetFormulas()
        {
            int days_number = (this.WorksEndDate - this.RecordCardStartDate).Days;

            Excel.Range tmp_first_rc_card_days_row = null;
            if (this.Owner == null && this.WorksSections.Count > 0
                && this.WorksSections[0].MSGWorks.Count > 0
                && this.WorksSections[0].MSGWorks[0].VOVRWorks.Count > 0
                && this.WorksSections[0].MSGWorks[0].VOVRWorks[0].KSWorks.Count > 0
                && this.WorksSections[0].MSGWorks[0].VOVRWorks[0].KSWorks[0].RCWorks.Count > 0)
            {
                RCWork first_rc_work = this.WorksSections[0].MSGWorks[0].VOVRWorks[0].KSWorks[0].RCWorks[0];
                Excel.Range first_cell = this.RegisterSheet.Cells[first_rc_work.GetTopRow(), WorkReportCard.WRC_PC_QUANTITY_COL];
                Excel.Range last_cell = this.RegisterSheet.Cells[first_rc_work.GetTopRow(), WorkReportCard.WRC_PC_QUANTITY_COL + days_number + 1];

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
                tmp_first_rc_card_days_row = this.RegisterSheet.Range[first_cell, this.RegisterSheet.Cells[first_cell.Row, first_cell.Column + date_iterator]];

            }

            foreach (WorksSection section in this.WorksSections)
            {
                foreach (MSGWork msg_work in section.MSGWorks)
                {
                    string msg_works_labourness_sum_formula = "";
                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        string vovr_works_labourness_sum_formula = "";
                        foreach (KSWork ks_work in vovr_work.KSWorks)
                        {
                            string rc_works_labourness_sum_formula = "";
                            if (this.Owner == null && tmp_first_rc_card_days_row != null) tmp_first_rc_card_days_row.Copy();
                            foreach (RCWork rc_work in ks_work.RCWorks)
                            {
                                if (rc_work.ReportCard == null)
                                {
                                    rc_work.ReportCard = new WorkReportCard();
                                    rc_work.ReportCard.Worksheet = rc_work.Worksheet;
                                    this.Register(rc_work.ReportCard, "Number", rc_work.GetTopRow(), WorkReportCard.WRC_NUMBER_COL, this.RegisterSheet);
                                    this.Register(rc_work.ReportCard, "PreviousComplatedQuantity", rc_work.GetTopRow(), WorkReportCard.WRC_PC_QUANTITY_COL, this.RegisterSheet);
                                    rc_work.ReportCard.Number = rc_work.Number;
                                    if (this.Owner == null) tmp_first_rc_card_days_row.Copy();

                                }

                                var first_cell = this.RegisterSheet.Cells[rc_work.GetTopRow(), WorkReportCard.WRC_PC_QUANTITY_COL];
                                var lastt_cell = this.RegisterSheet.Cells[rc_work.GetTopRow(), WorkReportCard.WRC_PC_QUANTITY_COL + 1 + days_number];
                                Excel.Range q_summ_range = this.RegisterSheet.Cells[rc_work.GetTopRow(), RCWork.RC_QUANTITY_FACT_COL];
                                q_summ_range.Formula = $"=SUM({Func.RangeAddress(first_cell)}:{Func.RangeAddress(lastt_cell)})";

                                if (this.Owner == null)
                                {
                                    Excel.Range w_days_row_range = this.RegisterSheet.Cells[rc_work.ReportCard.GetTopRow(), WorkReportCard.WRC_PC_QUANTITY_COL];
                                    if (tmp_first_rc_card_days_row != null && this.Owner == null)
                                    {
                                        //        tmp_first_rc_card_days_row.Copy();
                                        w_days_row_range.PasteSpecial(XlPasteType.xlPasteAll);
                                    }
                                }
                                rc_works_labourness_sum_formula +=
                                       $"{Func.RangeAddress(rc_work["Quantity"].Cell)}*{Func.RangeAddress(rc_work["Laboriousness"].Cell)}+";


                            }
                            rc_works_labourness_sum_formula = rc_works_labourness_sum_formula.TrimEnd('+');
                            if (rc_works_labourness_sum_formula != "")
                            {
                                string ks_quantity_formula = $"=({rc_works_labourness_sum_formula})/{Func.RangeAddress(ks_work["Laboriousness"].Cell)}";
                                ks_work["Quantity"].Cell.Formula = ks_quantity_formula;

                                vovr_works_labourness_sum_formula +=
                                    $"{Func.RangeAddress(ks_work["Quantity"].Cell)}*{Func.RangeAddress(ks_work["Laboriousness"].Cell)}+";
                            }
                        }
                        vovr_works_labourness_sum_formula = vovr_works_labourness_sum_formula.TrimEnd('+');

                        if (vovr_works_labourness_sum_formula != "")
                        {
                            string vovr_quantity_formula = $"=({vovr_works_labourness_sum_formula})/{Func.RangeAddress(vovr_work["Laboriousness"].Cell)}";
                            vovr_work["Quantity"].Cell.Formula = vovr_quantity_formula;
                        }
                        msg_works_labourness_sum_formula +=
                                             $"{Func.RangeAddress(vovr_work["Quantity"].Cell)}*{Func.RangeAddress(vovr_work["Laboriousness"].Cell)}+";

                    }
                    msg_works_labourness_sum_formula = msg_works_labourness_sum_formula.TrimEnd('+');
                    if (msg_works_labourness_sum_formula != "")
                    {
                        string msg_quantity_formula = $"=({msg_works_labourness_sum_formula})/{Func.RangeAddress(msg_work["Laboriousness"].Cell)}";
                        msg_work["Quantity"].Cell.Formula = msg_quantity_formula;
                    }

                }

            }

            foreach (WorkerConsumption consumption in this.WorkerConsumptions)
            {
                int col_iterator = WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL;
                while (col_iterator <= (this.WorksEndDate - this.RecordCardStartDate).Days)
                {
                    var cons_day_range = this.WorkerConsumptionsSheet.Cells[consumption["Number"].Row, col_iterator];
                    string cons_quantity_formula = "";
                    foreach (MSGExellModel model in this.Children)
                    {
                        //var child_consumption = model.WorkerConsumptions.FirstOrDefault(cn => cn.Number == consumption.Number);
                        //if (child_consumption != null)
                        //{
                        //    int cons_row = child_consumption["Number"].Row;
                        //    var child_cons_day_range =
                        //         model.WorkerConsumptionsSheet.Cells[cons_row, col_iterator];
                        //    cons_quantity_formula += $"{model.WorkerConsumptionsSheet.Name}!{Func.RangeAddress(cons_day_range)}+";
                        //}
                        if (model.WorkerConsumptionsSheet != null)
                        {
                            int cons_row = consumption["Number"].Row;
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
                int col_iterator = MachineConsumption.MCH_CONSUMPTIONS_FIRST_DATE_COL;
                while (col_iterator <= (this.WorksEndDate - this.RecordCardStartDate).Days)
                {
                    var cons_day_range = this.MachineConsumptionsSheet.Cells[consumption["Number"].Row, col_iterator];
                    string cons_quantity_formula = "";
                    foreach (MSGExellModel model in this.Children)
                    {
                        //var child_consumption = model.MachineConsumptions.FirstOrDefault(cn => cn.Number == consumption.Number);
                        //if (child_consumption != null)
                        //{
                        //    int cons_row = child_consumption["Number"].Row;
                        //    var child_cons_day_range =
                        //         model.MachineConsumptionsSheet.Cells[cons_row, col_iterator];
                        //    cons_quantity_formula += $"{model.MachineConsumptionsSheet.Name}!{Func.RangeAddress(cons_day_range)}+";
                        //}
                        if (model.MachineConsumptionsSheet != null)
                        {
                            int cons_row = consumption["Number"].Row;
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
        /// Функиця пересчета трудоемкостей всех типов работ исходя из проставленных в трудоемкостей
        /// в работах типа КС-2
        /// </summary>
        public void CalcLabourness()
        {
            foreach (WorksSection section in this.WorksSections)
            {
                foreach (MSGWork msg_work in section.MSGWorks)
                {
                    if (msg_work.Number == "1.13")
                        ;
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
                        if (common_vovr_laboueness != 0)
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

                    rc_work[nameof(rc_work.LabournessCoefficient)].IsValid = false;
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
                    if (msg_work.ReportCard == null)
                    {
                        msg_work.ReportCard = new WorkReportCard();
                        msg_work.ReportCard.Number = msg_work.Number;
                    }
                    else
                        msg_work.ReportCard.Clear();

                    var msg_work_all_rcWorks = this.RCWorks.Where(w => w.Number.StartsWith(msg_work.Number + "."));
                    decimal pr_works_loboriosness_summ = 0;
                    foreach (RCWork rc_work in msg_work_all_rcWorks)
                    {
                        if (rc_work.ReportCard != null)
                        {
                            foreach (WorkDay rc_w_day in rc_work.ReportCard)
                            {
                                rc_w_day.LaborСosts = rc_w_day.Quantity * rc_work.Laboriousness;

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
                            rc_work.PreviousComplatedQuantity = rc_work.ReportCard.PreviousComplatedQuantity;
                            pr_works_loboriosness_summ += rc_work.PreviousComplatedQuantity * rc_work.Laboriousness;

                        }
                        msg_work.PreviousComplatedQuantity = pr_works_loboriosness_summ / msg_work.Laboriousness;
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
            int rowIndex = WorkerConsumption.W_CONSUMPTIONS_FIRST_ROW_INDEX;
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

                                    //foreach (KeyValuePair<string, ExcelPropAddress> map_item in child_w_day.CellAddressesMap)
                                    //{
                                    //    int date_index = 0;
                                    //    while (this.WorkerConsumptionsSheet.Cells[WorkerConsumption.W_CONSUMPTIONS_DATE_RAW, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value != null &&
                                    //      DateTime.Parse(this.WorkerConsumptionsSheet.Cells[WorkerConsumption.W_CONSUMPTIONS_DATE_RAW, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString()) < end_date)
                                    //    {
                                    //        if (DateTime.Parse(this.WorkerConsumptionsSheet.Cells[WorkerConsumption.W_CONSUMPTIONS_DATE_RAW, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString()) == curent_w_day.Date)
                                    //            break;
                                    //        date_index++;
                                    //    }
                                    //    int curent_w_consumption_row = worker_consumption.WorkersConsumptionReportCard.GetPropAddress("WorkersConsumptionReportCard").Row;

                                    //    this.Register(curent_w_day,map_item.Key, new ExcelPropAddress(curent_w_consumption_row, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + date_index, this.WorkerConsumptionsSheet, typeof(DateTime)));

                                    //    curent_w_day.Quantity = curent_w_day.Quantity;
                                    //    //          this.WorkerConsumptionsSheet.Cells[curent_w_consumption_row, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + date_index] =
                                    //    //      curent_w_day.Quantity.ToString();

                                    //    //   this.Register(curent_w_day);
                                    //}
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
        }
        public void UpdateExcelReportCardsRepresetation(IEnumerable<Tuple<string, int>> numbers_rows = null)
        {
            if (numbers_rows != null)
            {

                foreach (Tuple<string, int> tuple in numbers_rows)
                {
                    var report_card = this.WorkReportCards.FirstOrDefault(rc => rc.Number == tuple.Item1);
                    if (report_card != null)
                    {
                        report_card.DaysFirsDate = this.RecordCardStartDate;

                        report_card.AdjustExcelRepresentionTree(tuple.Item2);
                        report_card.UpdateExcelRepresetation();
                    }
                }
                return;
            }
            foreach (RCWork rc_work in this.RCWorks)
            {
                var report_card = rc_work.ReportCard;
                if (report_card != null)
                {
                    report_card.AdjustExcelRepresentionTree(rc_work.GetTopRow());
                    report_card.UpdateExcelRepresetation();
                }
            }
        }


        /// <summary>
        /// 
        /// или если ведомость сама общая, то просто очищает у нее каледарную часть с записями выполенных объемов
        /// </summary>
        public override void UpdateExcelRepresetation()
        {
            //this.WorksSections.Worksheet = this.RegisterSheet;
            //if(this.Owner!=null)
            //{

            //}
            //  this.ClearWorksheetCommonPart();
            if (this.WorksSections.Count == 0) return;

            int last_row = FIRST_ROW_INDEX - _SECTIONS_GAP;
            foreach (WorksSection w_section in this.WorksSections.OrderBy(s => Int32.Parse(s.Number)))
            {
                last_row = w_section.AdjustExcelRepresentionTree(last_row + _SECTIONS_GAP);
                w_section.UpdateExcelRepresetation();
            }

            foreach(var rc in this.WorkReportCards)
            {
                this.UpdateExcelReportCardsRepresetation(this.RCWorks.Select(s => new Tuple<string, int>(s.Number, s.GetTopRow())));
            }

            Excel.Range all_sections_lowest_range = this.WorksSections.GetRange()?.GetRangeWithLowestEdge();
            int lowest_row = all_sections_lowest_range.Rows[all_sections_lowest_range.Rows.Count].Row;

            int section_first_row = lowest_row + _SECTIONS_GAP_FOR_INVALID_OBJECTS - _SECTIONS_GAP;
            int msg_work_first_row = lowest_row + _SECTIONS_GAP_FOR_INVALID_OBJECTS - _MSG_WORKS_GAP;
            int vovr_work_first_row = lowest_row + _SECTIONS_GAP_FOR_INVALID_OBJECTS;
            int ks_work_first_row = lowest_row + _SECTIONS_GAP_FOR_INVALID_OBJECTS;
            int rc_work_first_row = lowest_row + _SECTIONS_GAP_FOR_INVALID_OBJECTS;
            int repordCard_first_row = lowest_row + _SECTIONS_GAP_FOR_INVALID_OBJECTS;

            foreach (IObservableExcelBindableBase obj in this.InvalidObjects)
            {

                if (obj is WorksSection section) section_first_row = section.AdjustExcelRepresentionTree(section_first_row + _SECTIONS_GAP);
                if (obj is MSGWork msg_work) msg_work_first_row = msg_work.AdjustExcelRepresentionTree(msg_work_first_row + _MSG_WORKS_GAP);
                if (obj is VOVRWork vovr_work) vovr_work_first_row = vovr_work.AdjustExcelRepresentionTree(vovr_work_first_row);
                if (obj is KSWork ks_work) ks_work_first_row = ks_work.AdjustExcelRepresentionTree(ks_work_first_row);
                if (obj is RCWork rc_work) rc_work_first_row = rc_work.AdjustExcelRepresentionTree(rc_work_first_row);
                if (obj is WorkReportCard reportCard) repordCard_first_row = reportCard.AdjustExcelRepresentionTree(repordCard_first_row++);
                obj.UpdateExcelRepresetation();
            }

          

        }

        public  void UpdateExcelWorkerConsuptionsRepresentation()
        {
            foreach (var wc in this.WorkerConsumptions)
            {
                foreach (var w in wc.WorkersConsumptionReportCard)
                {
                    int d_col = (w.Date - this.RecordCardStartDate).Days;
                    w.ChangeLeftColumn(WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + d_col);
                    w.UpdateExcelRepresetation();
                }

            }
        }
        public void UpdateExcelMachineConsuptionsRepresentation()
        {
            foreach (var mc in this.MachineConsumptions)
            {
                foreach (var m in mc.MachinesConsumptionReportCard)
                {
                    int d_col = (m.Date - this.RecordCardStartDate).Days;
                    m.ChangeLeftColumn(WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + d_col);
                    m.UpdateExcelRepresetation();
                }

            }
        }

        public void UpdateRecordCardsArea()
        {
            if (this.Owner != null)
            {
                DateTime date = this.RecordCardStartDate;
                this.ClearAllRecordCards();
                this.LoadWorksReportCards();
                this.RecordCardStartDate = this.Owner.RecordCardStartDate;
                this.ClearWorksheetRecorCardPart(); ;
                this.UpdateExcelReportCardsRepresetation(this.Owner.RCWorks.Select(s => new Tuple<string, int>(s.Number, s.GetTopRow())));
                this.RecordCardStartDate = date;
               }

        }
        public void UpdateReegisterSheet()
        {
            if (this.Owner != null)
            {

                this.ClearWorksheetCommonPart();
                this.Owner.WorksSections.Worksheet = this.RegisterSheet;
                this.Owner.UpdateExcelRepresetation();
                this.Owner.WorksSections.Worksheet = this.Owner.RegisterSheet;
             //   this.UpdateExcelReportCardsRepresetation(this.Owner.RCWorks.Select(s => new Tuple<string, int>(s.Number, s.GetTopRow())));
             
            }

        }
        public void UpdateWorkerConsumptionsArea()
        {
            if (this.Owner != null)
            {
                DateTime date = this.RecordCardStartDate;
                this.ClearAllWorkerConsuptions();
                this.LoadWorkerConsumptions();
                this.RecordCardStartDate = this.Owner.RecordCardStartDate;
                this.ClearWorksheetWorkerConsumptionsPart();
                this.UpdateExcelWorkerConsuptionsRepresentation();
                 this.RecordCardStartDate = date;
            }

        }
        public void UpdateMachineConsumptionsArea()
        {
            if (this.Owner != null)
            {
                DateTime date = this.RecordCardStartDate;
                this.ClearAllMachineConsuptions();
                this.LoadMachineConsumptions();
                this.RecordCardStartDate = this.Owner.RecordCardStartDate;
                this.ClearWorksheetMachineConsumptionsPart();
                this.UpdateExcelMachineConsuptionsRepresentation();
                this.RecordCardStartDate = date;
            }

        }
        public void UpdateAll()
        {
            {
                if (this.Owner != null)
                {
                    this.WorkedDaysNumber = this.Owner.WorkedDaysNumber;
                    this.UpdateReegisterSheet();
                    this.UpdateRecordCardsArea();
                    this.UpdateWorkerConsumptionsArea();
                    this.UpdateMachineConsumptionsArea();

                    this.RecordCardStartDate = this.Owner.RecordCardStartDate;
                    this.RegisterSheet.Cells[MSGExellModel.WORKS_START_DATE_ROW, MSGExellModel.WORKS_START_DATE_COL].Value = this.RecordCardStartDate.ToString("d");
                    this.WorkerConsumptionsSheet.Cells[WorkerConsumption.W_CONSUMPTIONS_DATE_RAW, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL].Value = this.RecordCardStartDate.ToString("d");
                    this.MachineConsumptionsSheet.Cells[MachineConsumption.MCH_CONSUMPTIONS_DATE_RAW, MachineConsumption.MCH_CONSUMPTIONS_FIRST_DATE_COL].Value = this.RecordCardStartDate.ToString("d");

                    this.ReloadSheetModel();
                    this.SetFormulas();
                    this.LoadMSGWorks();
                    this.SetStyleFormats();


                }
                else
                {
                    this.ClearWorksheetCommonPart();
                    this.ClearWorksheetRecorCardPart();

                    this.UpdateExcelRepresetation();
                    this.SetFormulas();
                    this.SetStyleFormats();
                }

            }
        }


        /// <summary>
        /// Функция копирует часть объектой модели из родительской модеи в текущую
        /// </summary>
        public void CopyOwnerObjectModels()
        {
            if (this.Owner != null)
            {
                this.WorksStartDate = this.Owner.WorksStartDate;

                this.Unregister(this.WorksSections);

                this.WorksSections = (ExcelNotifyChangedCollection<WorksSection>)this.Owner.WorksSections.Clone();
                this.WorksSections.Owner = this;
                this.SetCommonModelCollections();


                //    foreach (var section in this.WorksSections)
                //      this.RegisterObjectInObjectPropertyNameRegister(section);
            }

        }
        /// <summary>
        /// Функция заполняет соосветврующие общие коллекции из дерева загруженных в объекты данных
        /// </summary>
        public void SetCommonModelCollections()
        {
            this.WorksSections.Worksheet = this.RegisterSheet;
            this.MSGWorks.Worksheet = this.RegisterSheet;
            this.VOVRWorks.Worksheet = this.RegisterSheet;
            this.KSWorks.Worksheet = this.RegisterSheet;
            this.RCWorks.Worksheet = this.RegisterSheet;
            this.WorkReportCards.Worksheet = this.RegisterSheet;
            this.MSGWorks.Clear();
            this.VOVRWorks.Clear();
            this.KSWorks.Clear();
            this.RCWorks.Clear();
            foreach (WorksSection w_section in this.WorksSections)
            {
                w_section.Worksheet = this.RegisterSheet;
                foreach (MSGWork msg_work in w_section.MSGWorks)
                {
                    msg_work.Worksheet = this.RegisterSheet;
                    msg_work.Owner = w_section;
                    if (!this.MSGWorks.Contains(msg_work)) this.MSGWorks.Add(msg_work);
                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        vovr_work.Worksheet = this.RegisterSheet;
                        vovr_work.Owner = msg_work;
                        if (!this.VOVRWorks.Contains(vovr_work)) this.VOVRWorks.Add(vovr_work);

                        foreach (KSWork ks_work in vovr_work.KSWorks)
                        {
                            ks_work.Worksheet = this.RegisterSheet;
                            ks_work.Owner = vovr_work;
                            if (!this.KSWorks.Contains(ks_work)) this.KSWorks.Add(ks_work);
                            foreach (RCWork rc_work in ks_work.RCWorks)
                            {
                                rc_work.Worksheet = this.RegisterSheet;
                                rc_work.Owner = ks_work;
                                if (!this.RCWorks.Contains(rc_work)) this.RCWorks.Add(rc_work);
                            }
                        }
                    }
                }

            }
        }

        public void UpdateCellAddressMapsWorkSheets_bk()
        {

            this.WorksSections.Worksheet = this.RegisterSheet;
            foreach (WorksSection w_section in this.WorksSections)
            {
                w_section.Worksheet = this.RegisterSheet;
                foreach (MSGWork msg_work in w_section.MSGWorks)
                {
                    msg_work.Worksheet = this.RegisterSheet;
                    //   if (!this.MSGWorks.Contains(msg_work)) this.MSGWorks.Add(msg_work);

                    foreach (WorkScheduleChunk w_ch in msg_work.WorkSchedules)
                        w_ch.Worksheet = this.RegisterSheet;

                    foreach (NeedsOfWorker n_w in msg_work.WorkersComposition)
                        n_w.Worksheet = this.RegisterSheet;

                    foreach (NeedsOfMachine n_m in msg_work.MachinesComposition)
                        n_m.Worksheet = this.RegisterSheet;

                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        //  if (!this.VOVRWorks.Contains(vovr_work)) this.VOVRWorks.Add(vovr_work);
                        vovr_work.Worksheet = this.RegisterSheet;
                        foreach (KSWork ks_work in vovr_work.KSWorks)
                        {
                            //      if (!this.KSWorks.Contains(ks_work)) this.KSWorks.Add(ks_work);
                            ks_work.Worksheet = this.RegisterSheet;
                            foreach (RCWork rc_work in ks_work.RCWorks)
                            {
                                //    if (!this.RCWorks.Contains(rc_work)) this.RCWorks.Add(rc_work);
                                rc_work.ReportCard = this.WorkReportCards.Where(rc => rc.Number == rc_work.Number).FirstOrDefault();
                                rc_work.Worksheet = this.RegisterSheet;
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
            Excel.Range common_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, WorkReportCard.WRC_DATE_COL],
                this.RegisterSheet.Cells[this.KSWorks[this.KSWorks.Count - 1].GetPropAddress("Laboriousness").Row,
                                          this.KSWorks[this.KSWorks.Count - 1].GetPropAddress("Laboriousness").Column]];
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
                common_area_range = this.WorkerConsumptionsSheet.Range[this.WorkerConsumptionsSheet.Cells[WorkerConsumption.W_CONSUMPTIONS_FIRST_ROW_INDEX, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL],
                            this.WorkerConsumptionsSheet.Cells[last_work_consumption_Day["Quantity"].Row, last_work_consumption_Day["Quantity"].Column]];
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

                Excel.Range common_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, WorksSection.WSEC_NUMBER_COL],
                      this.RegisterSheet.Cells[last_cell.Row, WorkReportCard.WRC_NUMBER_COL - 1]];
                common_area_range.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                common_area_range.Interior.ColorIndex = 0;

                common_area_range.ClearContents();

            }
            catch (Exception exp)
            {
                throw new Exception($"Ошибка при очистке листа.Ошибка:{exp.Message}");
            }

            this.RemoveGroups(this.RegisterSheet);

        }

        public void ClearWorksheetRecorCardPart()
        {
            try
            {
                Excel.Range last_cell = this.RegisterSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

                Excel.Range record_cards_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, WorkReportCard.WRC_NUMBER_COL],
                    this.RegisterSheet.Cells[last_cell.Row, last_cell.Column]];
                record_cards_area_range.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                record_cards_area_range.Interior.ColorIndex = 0;

                record_cards_area_range.ClearContents();
            }
            catch (Exception exp)
            {
                throw new Exception($"Ошибка при очистке листа.Ошибка:{exp.Message}");
            }

            this.RemoveGroups(this.RegisterSheet);

        }
        public void ClearWorksheetWorkerConsumptionsPart()
        {
            try
            {
                Excel.Range last_cell = this.WorkerConsumptionsSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

                Excel.Range common_area_range = this.WorkerConsumptionsSheet.Range[
                    this.WorkerConsumptionsSheet.Cells[WorkerConsumption.W_CONSUMPTIONS_FIRST_ROW_INDEX, WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL],
                      this.WorkerConsumptionsSheet.Cells[last_cell.Row, last_cell.Column]];
                common_area_range.ClearContents();

            }
            catch (Exception exp)
            {
                throw new Exception($"Ошибка при очистке листа.Ошибка:{exp.Message}");
            }

   
        }
        public void ClearWorksheetMachineConsumptionsPart()
        {
            try
            {
                Excel.Range last_cell = this.MachineConsumptionsSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

                Excel.Range common_area_range = this.MachineConsumptionsSheet.Range[
                    this.MachineConsumptionsSheet.Cells[MachineConsumption.MCH_CONSUMPTIONS_FIRST_ROW_INDEX, MachineConsumption.MCH_CONSUMPTIONS_FIRST_DATE_COL],
                      this.MachineConsumptionsSheet.Cells[last_cell.Row, last_cell.Column]];
                common_area_range.ClearContents();

            }
            catch (Exception exp)
            {
                throw new Exception($"Ошибка при очистке листа.Ошибка:{exp.Message}");
            }


        }
        public void ClearWorksheetAll()
        {
            this.ClearWorksheetCommonPart();
            this.ClearWorksheetRecorCardPart();
        }
        /// <summary>
        /// Удалфет все групы в сторках и столбцах
        /// </summary>
        /// <param name="worksheet"></param>
        public void RemoveGroups(Excel.Worksheet worksheet)
        {
            Excel.Range all_rows = worksheet.Cells.Rows;
            Excel.Range all_colomns = worksheet.Cells;



            for (int ii = 0; ii < 5; ii++)
                try
                {
                    //    all_rows.Select();
                    all_rows.Ungroup();
                    //    all_colomns.Select();
                    all_colomns.Ungroup();

                }
                catch (Exception exp)
                {
                    //      throw new Exception($"Ошибка при удалении всех группировок. Ошибка:{exp.Message}");
                }
        }
        /// <summary>
        /// Функция получает ближайший на листе Exсуд объет необходимого типа.
        /// </summary>
        /// <param name="section"></param>
        /// <param name="object_type"></param>
        /// <returns></returns>
        public List<IObservableExcelBindableBase> GetObjectsBySelection(Excel.Range selection, Type object_type)
        {
            ObservableCollection<Tuple<double, IObservableExcelBindableBase>> objects_distation = new ObservableCollection<Tuple<double, IObservableExcelBindableBase>>();
            List<IObservableExcelBindableBase> finded_objects = new List<IObservableExcelBindableBase>();
            int top_row = selection.Rows[1].Row;
            int bottom_row = selection.Rows[selection.Rows.Count].Row;
            List<IObservableExcelBindableBase> uniq_objcts = new List<IObservableExcelBindableBase>();

            foreach (var subscr_obj in this.ExcelSubsriptions.Where(rr => rr.Observable is IObservableExcelBindableBase obj
                                                           && (obj.GetType() == object_type || obj.GetType().GetInterface(object_type.Name) != null)
                                                               && obj.Owner != null
                                                               && obj.GetTopRow() >= top_row
                                                               && obj.GetTopRow() <= bottom_row))
            {
                var obj = subscr_obj.Observable as IObservableExcelBindableBase;
                int obj_row = obj.GetTopRow();
                int obj_col = obj.GetLeftColumn();
                double dist = Math.Sqrt(Math.Pow(obj_row - selection.Row, 2) + Math.Pow(obj_col - selection.Column, 2));
                if (uniq_objcts.FirstOrDefault(ob => ob.Id == obj.Id) == null)
                {
                    objects_distation.Add(new Tuple<double, IObservableExcelBindableBase>(dist, obj));
                    uniq_objcts.Add(obj);
                }

            }
            IObservableExcelBindableBase finded_obj = null;
            var tuple = objects_distation.OrderBy(el => el.Item1).FirstOrDefault();
            if (tuple != null)
                finded_obj = tuple.Item2 as IObservableExcelBindableBase;

            foreach (var kvp in objects_distation.OrderBy(_kvp => _kvp.Item1))
                finded_objects.Add(kvp.Item2);

            if (selection.Rows.Count == 1)
                return new List<IObservableExcelBindableBase>() { finded_obj };

            return finded_objects;
        }
        public List<IObservableExcelBindableBase> GetObjectsBySelection(Excel.Range selection, Func<IObservableExcelBindableBase, bool> obj_predicate)
        {
            ObservableCollection<Tuple<double, IObservableExcelBindableBase>> objects_distation = new ObservableCollection<Tuple<double, IObservableExcelBindableBase>>();
            List<IObservableExcelBindableBase> finded_objects = new List<IObservableExcelBindableBase>();
            int top_row = selection.Rows[1].Row;
            int bottom_row = selection.Rows[selection.Rows.Count].Row;
            List<IObservableExcelBindableBase> uniq_objcts = new List<IObservableExcelBindableBase>();

            foreach (var subscr_obj in this.ExcelSubsriptions.Where(rr => obj_predicate(rr.Observable as IObservableExcelBindableBase)))
            {
                var obj = subscr_obj.Observable as IObservableExcelBindableBase;
                int obj_row = obj.GetTopRow();
                int obj_col = obj.GetLeftColumn();
                double dist = Math.Sqrt(Math.Pow(obj_row - selection.Row, 2) + Math.Pow(obj_col - selection.Column, 2));
                if (!uniq_objcts.Contains(obj))
                {
                    objects_distation.Add(new Tuple<double, IObservableExcelBindableBase>(dist, obj));
                    uniq_objcts.Add(obj);
                }

            }
            IObservableExcelBindableBase finded_obj = null;
            var tuple = objects_distation.OrderBy(el => el.Item1).FirstOrDefault();
            if (tuple != null)
                finded_obj = tuple.Item2 as IObservableExcelBindableBase;

            foreach (var kvp in objects_distation.OrderBy(_kvp => _kvp.Item1))
                finded_objects.Add(kvp.Item2);

            return finded_objects;
        }

        public IObservableExcelBindableBase SelectedObjectForInsert(Excel.Range selection, Func<IObservableExcelBindableBase, bool> obj_predicate)
        {
            ObservableCollection<Tuple<double, IObservableExcelBindableBase>> objects_distation = new ObservableCollection<Tuple<double, IObservableExcelBindableBase>>();
            List<IObservableExcelBindableBase> finded_objects = new List<IObservableExcelBindableBase>();
            int top_row = selection.Rows[1].Row;
            int bottom_row = selection.Rows[selection.Rows.Count].Row;
            List<IObservableExcelBindableBase> uniq_objcts = new List<IObservableExcelBindableBase>();

            foreach (var subscr_obj in this.ExcelSubsriptions.Where(rr => obj_predicate(rr.Observable as IObservableExcelBindableBase)))
            {
                var obj = subscr_obj.Observable as IObservableExcelBindableBase;
                int obj_row = obj.GetTopRow();
                int obj_col = obj.GetLeftColumn();
                double dist = Math.Sqrt(Math.Pow(obj_row - selection.Row, 2) + Math.Pow(obj_col - selection.Column, 2));
                if (!uniq_objcts.Contains(obj))
                {
                    objects_distation.Add(new Tuple<double, IObservableExcelBindableBase>(dist, obj));
                    uniq_objcts.Add(obj);
                }

            }
            IObservableExcelBindableBase finded_obj = null;
            var tuple = objects_distation.OrderBy(el => el.Item1).FirstOrDefault();
            foreach (var kvp in objects_distation.OrderBy(_kvp => _kvp.Item1))
                finded_objects.Add(kvp.Item2);

            var above_obj = finded_objects.FirstOrDefault(ob => ob.GetTopRow() >= selection.Row);

            if (above_obj != null)
                finded_obj = above_obj;

            return finded_obj;
        }

        public void InsertRow(int row)
        {


        }
    }
}
