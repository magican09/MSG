﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public class WorkReportCard: ExcelNotifyChangedCollection<WorkDay>
    {
     
        private string _number;

        public string Number
        {
            get { return _number; }
            set { SetProperty(ref _number, value); }
        }//Номер работы

        private decimal _quantity;

        public decimal Quantity
        {
            get {
                decimal out_value = 0;
                foreach (WorkDay work_day in this)
                    out_value += work_day.Quantity;
                _quantity = out_value;
                return _quantity; }
            
        }//Выполенный объем работ
        
    }
}
