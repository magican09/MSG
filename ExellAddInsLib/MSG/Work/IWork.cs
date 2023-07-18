namespace ExellAddInsLib.MSG
{
    public interface IWork
    {
        System.Collections.ObjectModel.ObservableCollection<IWork> Children { get; set; }
        decimal Laboriousness { get; set; }
        string Name { get; set; }
        string Number { get; set; }
        IWork Owner { get; set; }
        MSGExellModel OwnerExellModel { get; set; }
        decimal ProjectQuantity { get; set; }
        decimal Quantity { get; set; }
        WorkReportCard ReportCard { get; set; }
        int RowIndex { get; set; }
        UnitOfMeasurement UnitOfMeasurement { get; set; }
    }
}