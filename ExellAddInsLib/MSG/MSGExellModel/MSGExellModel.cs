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
        //public const int WORKS_END_DATE_ROW = 2;
        //public const int WORKS_END_DATE_COL = 3;

        //public const int FIRST_ROW_INDEX = 7;
        //public const int MSG_NUMBER_COL = 2;
        //public const int MSG_NAME_COL = 3;
        //public const int MSG_MEASURE_COL = 4;
        //public const int MSG_QUANTITY_COL = 5;
        //public const int MSG_QUANTITY_FACT_COL = 6;
        //public const int MSG_LABOURNESS_COL = 7;
        //public const int MSG_START_DATE_COL = 8;
        //public const int MSG_END_DATE_COL = 9;


        //public const int VOVR_NUMBER_COL = 10;
        //public const int VOVR_NAME_COL = 11;
        //public const int VOVR_MEASURE_COL = 12;
        //public const int VOVR_QUANTITY_COL = 13;
        //public const int VOVR_QUANTITY_FACT_COL = 14;
        //public const int VOVR_LABOURNESS_COL = 15;


        //public const int KS_NUMBER_COL = 16;
        //public const int KS_CODE_COL = 17;
        //public const int KS_NAME_COL = 18;
        //public const int KS_MEASURE_COL = 19;
        //public const int KS_QUANTITY_COL = 20;
        //public const int KS_QUANTITY_FACT_COL = 21;
        //public const int KS_LABOURNESS_COL = 22;

        //public const int WRC_DATE_ROW = 6;
        //public const int WRC_NUMBER_COL = 23;
        //public const int WRC_DATE_COL = 24;

       
        public  MSGWorksCollection MSGWorks { get; private set; } = new MSGWorksCollection();
        public  ObservableCollection<VOVRWork> VOVRWorks { get; private set; } = new ObservableCollection<VOVRWork>();
        public  ObservableCollection<KSWork> KSWorks { get; private set; } = new ObservableCollection<KSWork>();

        
        public   Excel.Worksheet RegisterSheet { get; set; }
       
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
                   
            }
        }
        private void OnPropertyChange(object sender, PropertyChangedEventArgs e)
        {
           if(sender is ExcelBindableBase bindable_object)
            {
                RegisterSheet.Cells[bindable_object.CellAddressesMap[e.PropertyName].Item1,
                    bindable_object.CellAddressesMap[e.PropertyName].Item2] =   sender.GetType().GetProperty(e.PropertyName).GetValue(sender).ToString();
            }
          

        }

      

    }
}
