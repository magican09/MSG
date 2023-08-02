using System.Collections.ObjectModel;

namespace ExellAddInsLib.MSG
{
    public interface IWork : IExcelBindableBase
    {
        ObservableCollection<IWork> Children { get; set; }
        decimal Laboriousness { get; set; }
        string Name { get; set; }
        string Number { get; set; }
      //  IWork Owner { get; set; }
        MSGExellModel OwnerExellModel { get; set; }
        decimal ProjectQuantity { get; set; }
        decimal Quantity { get; set; }
        WorkReportCard ReportCard { get; set; }
        int RowIndex { get; set; }
        UnitOfMeasurement UnitOfMeasurement { get; set; }

    }
}