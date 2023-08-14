using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class MSGWork : Work
    {
        public const int MSG_NUMBER_COL = 4;
        public const int MSG_NAME_COL = MSG_NUMBER_COL + 1;
        public const int MSG_MEASURE_COL = MSG_NUMBER_COL + 2;
        public const int MSG_QUANTITY_COL = MSG_NUMBER_COL + 3;
        public const int MSG_QUANTITY_FACT_COL = MSG_NUMBER_COL + 4;
        public const int MSG_LABOURNESS_COL = MSG_NUMBER_COL + 5;
        public const int MSG_START_DATE_COL = MSG_NUMBER_COL + 6;
        public const int MSG_END_DATE_COL = MSG_NUMBER_COL + 7;
        public const int MSG_SUNDAY_IS_VOCATION_COL = MSG_NUMBER_COL + 8;

        public const int _MSG_WORKS_GAP = 1;
        private Excel.Worksheet _worksheet;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public override Excel.Worksheet Worksheet
        {
            get { return _worksheet; }
            set
            {
                _worksheet = value;
                this.VOVRWorks.Worksheet = _worksheet;
                this.WorkersComposition.Worksheet = _worksheet;
                this.MachinesComposition.Worksheet = _worksheet;
                this.WorkSchedules.Worksheet = _worksheet;
                this.CellAddressesMap.SetWorksheet(_worksheet);
            }
        }
        private bool _isSundayVocation = true;

        public bool IsSundayVocation
        {
            get { return _isSundayVocation; }
            set { _isSundayVocation = value; }
        }

        private WorkSchedule _workSchedules = new WorkSchedule();


        public WorkSchedule WorkSchedules
        {
            get { return _workSchedules; }
            set { _workSchedules = value; }
        }

        private AdjustableCollection<VOVRWork> _vOVRWorks = new AdjustableCollection<VOVRWork>();

        [NonRegisterInUpCellAddresMap]
        public AdjustableCollection<VOVRWork> VOVRWorks
        {
            get { return _vOVRWorks; }
            set { _vOVRWorks = value; }
        }
        public int? GetShedulesAllDaysNumber()
        {


            if (WorkSchedules.Count > 0)
            {
                // var time_span = new TimeSpan(
                bool is_sunday_vocation = true;
                int? days_count = 0;
                foreach (WorkScheduleChunk chunk in this.WorkSchedules)
                {
                    if (chunk.IsSundayVacationDay == "Да")
                        is_sunday_vocation = true;
                    else
                        is_sunday_vocation = false;

                    int worked_day_number = 0;
                    for (DateTime date = chunk.StartTime; date <= chunk.EndTime; date = date.AddDays(1)) //Находим количество рабочих дней
                        if (is_sunday_vocation == false || date.DayOfWeek != DayOfWeek.Sunday)
                            worked_day_number++;

                    days_count += worked_day_number;// (chunk.EndTime - chunk.StartTime)?.Days;
                }
                //var time_span = WorkSchedules[WorkSchedules.Count - 1].EndTime - WorkSchedules[0].StartTime;
                return days_count;
            }
            else
                return 0;
        }
        public MSGWork() : base()
        {
            this.VOVRWorks.Owner = this;
        }
        public override void UpdateExcelRepresetation()
        {
            MSGWork msg_work = this;
            this.UpdateExellBindableObject();
            foreach (WorkScheduleChunk w_ch in msg_work.WorkSchedules)
                w_ch.UpdateExellBindableObject();
            foreach (NeedsOfWorker n_w in msg_work.WorkersComposition)
                n_w.UpdateExellBindableObject();
            foreach (NeedsOfMachine n_m in msg_work.MachinesComposition)
                n_m.UpdateExellBindableObject();
            foreach (VOVRWork vovr_work in msg_work.VOVRWorks.OrderBy(w => w.Number))
                vovr_work.UpdateExcelRepresetation();

        }

        public override int AdjustExcelRepresentionTree(int row)
        {
            MSGWork msg_work = this;
            int msg_row = row;
            int msg_lowest_row = 0;
            msg_work.ChangeTopRow(msg_row);
            int sh_ch_row_iterator = 0;
            foreach (WorkScheduleChunk w_ch in msg_work.WorkSchedules)
            {
                w_ch.AdjustExcelRepresentionTree(msg_row + sh_ch_row_iterator);
                sh_ch_row_iterator++;
            }
            int nw_row_iterator = 0;
            foreach (NeedsOfWorker n_w in msg_work.WorkersComposition)
            {
                n_w.AdjustExcelRepresentionTree(msg_row + nw_row_iterator);
                nw_row_iterator++;
            }
            int nm_row_iterator = 0;
            foreach (NeedsOfMachine n_m in msg_work.MachinesComposition)
            {
                n_m.AdjustExcelRepresentionTree(msg_row + nm_row_iterator);
                nm_row_iterator++;
            }
            if (msg_row + sh_ch_row_iterator > msg_lowest_row) msg_lowest_row = msg_row + sh_ch_row_iterator;
            if (msg_row + nw_row_iterator > msg_lowest_row) msg_lowest_row = msg_row + nw_row_iterator;
            if (msg_row + nm_row_iterator > msg_lowest_row) msg_lowest_row = msg_row + nm_row_iterator;
            int vovr_row = msg_row;
            foreach (VOVRWork vovr_work in msg_work.VOVRWorks.OrderBy(w => Int32.Parse(w.Number.Replace($"{w.NumberPrefix}.", ""))))
                vovr_row = vovr_work.AdjustExcelRepresentionTree(vovr_row); ;

            if (vovr_row < msg_lowest_row)
                msg_row = msg_lowest_row;
            else
                msg_row = vovr_row;

            return msg_row;
        }

        public override void SetStyleFormats(int col)
        {
            MSGWork msg_work = this;
            int msg_work_col = col;
            var msg_work_range = msg_work.GetRange(MSG_LABOURNESS_COL);
            msg_work_range.Interior.ColorIndex = msg_work_col;
            msg_work_range.SetBordersLine();
            int first_row = msg_work.GetTopRow();
            int last_row = msg_work.GetTopRow();

            msg_work.WorkersComposition.GetRange().SetBordersLine();
            int need_of_workers_count = 0;
            foreach (NeedsOfWorker need_of_worker in msg_work.WorkersComposition)
            {
                var need_of_worker_range = need_of_worker.GetRange();
                need_of_worker_range.Interior.ColorIndex = msg_work_col;
                need_of_workers_count++;
            }

            msg_work.MachinesComposition.GetRange().SetBordersLine();

            int need_of_machine_count = 0;
            foreach (NeedsOfMachine need_of_machine in msg_work.MachinesComposition)
            {
                var need_of_machine_range = need_of_machine.GetRange();
                need_of_machine_range.Interior.ColorIndex = msg_work_col;
                need_of_machine_count++;
            }

            msg_work.WorkSchedules.GetRange().SetBordersLine();
            int chunks_count = 0;
            foreach (WorkScheduleChunk chunk in msg_work.WorkSchedules)
            {
                var work_composition_range = chunk.GetRange();
                work_composition_range.Interior.ColorIndex = msg_work_col;
                chunks_count++;
            }
            if (msg_work.VOVRWorks.Count > 0)
            {
                Excel.Range _works_left_edge_range = msg_work.VOVRWorks.Worksheet.Range[msg_work.VOVRWorks[0].CellAddressesMap["Number"].Cell,
                                                                            msg_work.VOVRWorks[msg_work.VOVRWorks.Count - 1].CellAddressesMap["Number"].Cell];
                _works_left_edge_range.SetBordersLine(XlLineStyle.xlLineStyleNone, XlLineStyle.xlDashDot, XlLineStyle.xlLineStyleNone, XlLineStyle.xlLineStyleNone);
                int vovr_work_col = msg_work_col + 1;
                foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    vovr_work.SetStyleFormats(vovr_work_col++);
            }




            try
            {
                var msg_work_full_range = msg_work.GetRange();
                Excel.Range lowest_edge_range = msg_work_full_range.GetRangeWithLowestEdge();
                Excel.Range range = Worksheet.Range[Worksheet.Rows[msg_work_full_range.Row + 1], lowest_edge_range.Rows[lowest_edge_range.Rows.Count + _MSG_WORKS_GAP]];
                range.Group();

            }
            catch
            {

            }



        }

        public override Range GetRange()
        {
            Excel.Range base_range = base.GetRange();
            Excel.Range vovr_works_range = this.VOVRWorks.GetRange();
            Excel.Range w_schedules_works_range = this.WorkSchedules.GetRange();
            Excel.Range range = Worksheet.Application.Union(new List<Excel.Range>() { base_range, vovr_works_range, w_schedules_works_range });
            return range;
        }

        public override object Clone()
        {
            MSGWork new_obj = (MSGWork)base.Clone();
            new_obj.WorkSchedules = (WorkSchedule)this.WorkSchedules.Clone();
            new_obj.VOVRWorks = (AdjustableCollection<VOVRWork>)this.VOVRWorks.Clone();
            new_obj.WorkSchedules.Owner = new_obj;
            new_obj.VOVRWorks.Owner = new_obj;
            return new_obj;
        }
    }
}
