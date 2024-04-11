using System.Collections.Generic;

namespace DofusItemPriceExcelPj.Objects
{
    internal class PriceHistory
    {
        public string Label { get; set; } = "";
        public IList<PriceOccurence> PriceValues { get; set; } = new List<PriceOccurence>();

        public override string ToString()
        {
            return Label;
        }
    }
}
