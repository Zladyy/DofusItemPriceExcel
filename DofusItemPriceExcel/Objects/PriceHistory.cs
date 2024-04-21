using System;
using System.Collections.Generic;
using System.Linq;

namespace DofusItemPriceExcelPj.Objects
{
    internal class PriceHistory
    {
        public string Label { get; set; } = "";
        public IList<PriceOccurence> PriceValues { get; set; } = new List<PriceOccurence>();
        public int MinPrice
        {
            get
            {
                var min1Values = PriceValues.Where(x => x.Price1 != 0);
                var min10Values = PriceValues.Where(x => x.Price10 != 0);

                var min1HasValues = min1Values.Any();
                var min10HasValues = min10Values.Any();

                if(min1HasValues)
                {
                    var min1 = min1Values.Min(x => x.Price1);
                    if(min10HasValues)
                    {
                        var min10 = min10Values.Min(x => x.Price10On10);
                        //Both
                        return Math.Min(min1, min10);
                    }
                    else
                    {
                        //1
                        return min1;
                    }
                }
                else if(min10HasValues)
                {
                    var min10 = min10Values.Min(x => x.Price10On10);
                    //10
                    return min10;
                }
                else
                {
                    //None ?!
                    return 0;
                }
            }
        }

        public int MaxPrice
        {
            get
            {
                var max1Values = PriceValues.Where(x => x.Price1 != 0);
                var max10Values = PriceValues.Where(x => x.Price10 != 0);

                var max1HasValues = max1Values.Any();
                var max10HasValues = max10Values.Any();

                if(max1HasValues)
                {
                    var max1 = max1Values.Max(x => x.Price1);
                    if(max10HasValues)
                    {
                        var max10 = max10Values.Max(x => x.Price10On10);
                        //Both
                        return Math.Max(max1, max10);
                    }
                    else
                    {
                        //1
                        return max1;
                    }
                }
                else if(max10HasValues)
                {
                    var max10 = max10Values.Max(x => x.Price10On10);
                    //10
                    return max10;
                }
                else
                {
                    //None ?!
                    return 0;
                }
            }
        }

        //public int MaxPrice
        //{
        //    get
        //    {
        //        var max1 = PriceValues.Max(x => x.Price1);
        //        var max10 = PriceValues.Max(x => x.Price10On10);
        //        return max1 > max10 ? max1 : max10;
        //    }
        //}

        public override string ToString()
        {
            return Label;
        }
    }
}
