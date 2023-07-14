namespace ExellAddInsLib.MSG
{
    public interface IWork
    {
        decimal Laboriousness { get; set; }
        string Name { get; set; }
        string Number { get; set; }
        decimal ProjectQuantity { get; set; }
        decimal Quantity { get; set; }
        UnitOfMeasurement UnitOfMeasurement { get; set; }

    }
}