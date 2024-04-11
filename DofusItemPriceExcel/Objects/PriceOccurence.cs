using System;
using System.Globalization;

namespace DofusItemPriceExcelPj.Objects
{
    internal class PriceOccurence
    {
        public DateTime Date { get; set; }
        public int Price1 { get; set; }
        public int Price10 { get; set; }

        public override string ToString()
        {
            return $"{Date} / {Price1.ToString("N0", CultureInfo.CurrentCulture)} / {Price10.ToString("N0", CultureInfo.CurrentCulture)}";
        }
    }
}
