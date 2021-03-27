using System;

namespace ExcelCsharp
{
    public class Order
    {
        public Guid Id { get; set; }
        public DateTime Date { get; set; }
        public Decimal OrderValue { get; set; }
        public bool Shipped { get; set; }
    }
}