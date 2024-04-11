using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
