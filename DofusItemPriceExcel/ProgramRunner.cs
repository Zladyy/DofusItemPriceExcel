using DofusItemPriceExcelPj.Objects;
using ExcelDataReader;
using System.Text;
using MiniExcelLibs;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System;
using System.Diagnostics;

namespace DofusItemPriceExcelPj
{
    internal class ProgramRunner
    {
        readonly string _excelFilePath = @"C:\Users\User\OneDrive - SECIB\Personnel\Jeux\Dofus\Parchos.xlsx";

        public void Run()
        {
            //Read range for each scroll type
            //Add a new sheet to the excel document (edit or delete previously existing one)
            //Add a table for each scroll with a daily continuance
            //Aggregate previously read data from ranges per scroll type

            IList<PriceHistory> historyPerItem = GetPricesHistory();
            historyPerItem = AggregateData(historyPerItem);

            ManageAggregatedSheet(historyPerItem);
        }

        private IList<PriceHistory> GetPricesHistory()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using(var stream = File.Open(_excelFilePath, FileMode.Open, FileAccess.Read))
            {
                using(var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet();
                    var dataTable = dataSet.Tables[0];
                    var histories = dataTable.Rows[1].ItemArray.Where(x => !(x is DBNull)).Distinct().Select(x => new PriceHistory { Label = x?.ToString() ?? string.Empty }).ToList();

                    for(int i = 3; i < dataTable.Rows.Count; i++)
                    {
                        var row = dataTable.Rows[i];
                        for(int j = 1; j < row.ItemArray.Length; j += 4)
                        {
                            if(!(row.ItemArray[j] is DBNull))
                            {
                                var elem = histories[j / 4];
                                elem.PriceValues.Add(new PriceOccurence
                                {
                                    Date = (DateTime)(row.ItemArray[j] ?? DateTime.MinValue),
                                    Price1 = (int)(double)(row.ItemArray[j + 1] ?? 0),
                                    Price10 = (int)(double)(row.ItemArray[j + 2] ?? 0)
                                });
                            }
                        }
                    }
                    return histories;
                }
            }
        }

        private IList<PriceHistory> AggregateData(IList<PriceHistory> historyPerItem)
        {
            var result = new List<PriceHistory>();
            foreach(var history in historyPerItem)
            {
                var values = new List<PriceOccurence>();
                foreach(var price in history.PriceValues.GroupBy(x => x.Date.Date))
                {
                    var avgPrice1 = (int)Math.Round(price.Select(x => x.Price1).Average(), 0);
                    var avgPrice10 = (int)Math.Round(price.Select(x => x.Price10).Average(), 0);
                    values.Add(new PriceOccurence
                    {
                        Date = price.Key,
                        Price1 = avgPrice1,
                        Price10 = avgPrice10
                    });
                }
                result.Add(new PriceHistory
                {
                    Label = history.Label,
                    PriceValues = values
                });
            }
            result = FillMissingDailyValues(result);
            return result;
        }

        private List<PriceHistory> FillMissingDailyValues(List<PriceHistory> result)
        {
            foreach(var history in result)
            {
                for(int i = 1; i < history.PriceValues.Count; i++)
                {
                    var oldPriceDate = history.PriceValues[i-1].Date;
                    var priceDate = history.PriceValues[i].Date;
                    var dayDiff = (priceDate - oldPriceDate).TotalDays;
                    if(dayDiff > 1)
                    {
                        do
                        {
                            var newDate = oldPriceDate.AddDays(1);
                            history.PriceValues.Insert(i, new PriceOccurence
                            {
                                Date = newDate
                            });
                            dayDiff = (priceDate - newDate).TotalDays;
                            oldPriceDate = newDate;
                            i++;
                        } while(dayDiff > 1);
                    }
                }
            }

            return result;
        }

        private void ManageAggregatedSheet(IList<PriceHistory> historyPerItem)
        {
            //var reader = MiniExcel.
            var excelApp = new Application
            {
                Visible = true,
                WindowState = XlWindowState.xlMaximized,
                DisplayAlerts = false
            };
            var book = excelApp.Workbooks.Open(_excelFilePath);
            var sheets = book.Sheets;

            var hasAggregatedDataSheet = sheets.Count > 1;
            if(hasAggregatedDataSheet)
            {
                _Worksheet toDelete = sheets[2];
                var name = toDelete.Name;
                toDelete.Delete();
            }

            _Worksheet aggregatedSheet = book.Sheets.Add(After: book.ActiveSheet);

            aggregatedSheet.Name = "AggData";
            aggregatedSheet.Cells[1, "A"].ColumnWidth = 1;
            aggregatedSheet.Cells[1, "A"].RowHeight = 10;

            for(int i = 0; i < historyPerItem.Count(); i++)
            {
                var history = historyPerItem.ElementAt(i);
                aggregatedSheet.Cells[2, 2 + 4 * i] = history.Label;
                Range titleRange = aggregatedSheet.Range[aggregatedSheet.Cells[2, 2 + 4 * i], aggregatedSheet.Cells[2, 4 + 4 * i]];
                titleRange.Merge();
                titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;


                for(int j = 0; j < history.PriceValues.Count; j++)
                {
                    var price = history.PriceValues[j];
                    aggregatedSheet.Cells[3 + j, 2 + 4 * i] = price.Date;
                    if(price.Price1 != 0)
                    {
                        aggregatedSheet.Cells[3 + j, 3 + 4 * i] = price.Price1;
                    }
                    if(price.Price10 != 0)
                    {
                        aggregatedSheet.Cells[3 + j, 4 + 4 * i] = price.Price10;
                    }
                }

                Range datesRange = aggregatedSheet.Range[aggregatedSheet.Cells[3, 2 + 4 * i], aggregatedSheet.Cells[2 + history.PriceValues.Count, 2 + 4 * i]];
                datesRange.NumberFormat = "DD/MM/YYYY";

                Range wholeItemRange = aggregatedSheet.Range[aggregatedSheet.Cells[2, 2 + 4 * i], aggregatedSheet.Cells[2 + history.PriceValues.Count, 4 + 4 * i]];
                Borders borders1 = wholeItemRange.Borders;

                //All borders
                borders1.LineStyle = XlLineStyle.xlContinuous;
                borders1.Weight = 2d;

                //Thick borders around
                titleRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                wholeItemRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            }

            aggregatedSheet.Columns.AutoFit();
            aggregatedSheet.Activate();
        }
    }
}
