using ClosedXML.Attributes;

namespace WebApi.Models;

public class Order
{
    [XLColumn(Order = 1)]
    public long Id { get; set; }

    [XLColumn(Header = "Order Date", Order = 4)]
    public DateTime Date { get; set; }

    [XLColumn(Header = "Has Shipped", Order = 2)]
    public bool HasShipped { get; set; }

    [XLColumn(Order = 3)]
    public decimal Amount { get; set; }

    [XLColumn(Ignore = true)]
    public string UserEmail { get; set; }
}