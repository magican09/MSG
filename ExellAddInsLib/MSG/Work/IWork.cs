using Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public interface IWork:IExcelBindableBase
    {
        System.Collections.ObjectModel.ObservableCollection<IWork> Children { get; set; }
        decimal Laboriousness { get; set; }
        MachinesComposition MachinesComposition { get; set; }
        MSGExellModel OwnerExellModel { get; set; }
        decimal PreviousComplatedQuantity { get; set; }
        decimal ProjectQuantity { get; set; }
        decimal Quantity { get; set; }
        WorkReportCard ReportCard { get; set; }
        int RowIndex { get; set; }
        UnitOfMeasurement UnitOfMeasurement { get; set; }
        WorkersComposition WorkersComposition { get; set; }
        void ClearCalculatesFields();
        void SetSectionNumber(string section_number);
    }
}