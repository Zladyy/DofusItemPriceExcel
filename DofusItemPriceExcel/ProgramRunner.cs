using DofusItemPriceExcel.Objects;
using DofusItemPriceExcelPj.Objects;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace DofusItemPriceExcelPj
{
    public class ProgramRunner
    {
        private int buySellTresholdPercent = 10;
        private readonly int _colPerItem = 6;
        private int _colPerTab => _colPerItem + 1;
        private string _excelFilePath = "";

        public void Run(RunOptions options)
        {
            _excelFilePath = options.FilePath;
            if(options.BuySellTresholdPercent > 0)
            {
                buySellTresholdPercent = options.BuySellTresholdPercent;
            }
            //Task whenPriceLoaded = new Task<object>((object obj) =>
            //{
            //    IList<PriceHistory> historyPerItem = GetPricesHistoryFromExcelFile();
            //    historyPerItem = AggregateData(historyPerItem);
            //    WriteAggregatedSheetAndCharts(historyPerItem);
            //});
            //var a = Task.Run(GetPricesHistoryFromOnlineSource).Result;

            IList<PriceHistory> historyPerItem = GetPricesHistoryFromExcelFile();
            historyPerItem = AggregateData(historyPerItem);
            WriteAggregatedSheetAndCharts(historyPerItem);
        }

        private async Task<object> GetPricesHistoryFromOnlineSource()
        {
            try
            {
                //TODO: Get price history from source :
                //https://www.vulbis.com/?server=Draconiros&gids=&percent=0&craftableonly=false&select-type=76&sellchoice=false&buyqty=1&sellqty=1&percentsell=0
                var client = new HttpClient();
                var msg = new HttpRequestMessage()
                {
                    Method = HttpMethod.Get,
                    RequestUri = new Uri("https://www.vulbis.com/?server=Draconiros&gids=&percent=0&craftableonly=false&select-type=76&sellchoice=false&buyqty=1&sellqty=1&percentsell=0"),
                    Headers =
                    {
                        Referrer = new Uri("https://www.vulbis.com"),

                    },
                };
                msg.Headers.UserAgent.Clear();

                msg.Headers.Add("cf_clearance", "wgr9ErXKg9O9.Jzawy8o1llo.ISxlmI1lNR3xrvQ0hE-1713397708-1.0.1.1-aZ1_6c.rUHCjeiLh17I.ooQ265kAoDWOFfyLgEriLFmFFADAH_Kg7_r3qvWAZFs47FgMG12jCvz0ips3gCTHBw");
                msg.Headers.Add("XeloriumServer", "Orukam");
                msg.Headers.Add("SrambadServer", "Orukam");
                msg.Headers.Add("EcaflipusServer", "Orukam");
                msg.Headers.Add("EnutrosorServer", "Orukam");
                msg.Headers.Add("BUY_QTY", "1");
                msg.Headers.Add("SELL_QTY", "1");
                msg.Headers.Add("PERCENT_SELL_CHOICE", "0");
                msg.Headers.Add("SERVER_CHOICE", "Draconiros");
                msg.Headers.Add("PERCENT_CHOICE", "0");
                msg.Headers.Add("TYPE", "76");
                var response = await client.SendAsync(msg);
                var str = await response.Content.ReadAsStringAsync();
            }
            catch
            {

            }



            return null;
        }

        private IList<PriceHistory> GetPricesHistoryFromExcelFile()
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
                                var valPrice1 = row.ItemArray[j + 1];
                                var valPrice10 = row.ItemArray[j + 2];
                                elem.PriceValues.Add(new PriceOccurence
                                {
                                    Date = (DateTime)(row.ItemArray[j] ?? DateTime.MinValue),
                                    Price1 = (int)(double)((valPrice1 is DBNull) ? 0d : valPrice1),
                                    Price10 = (int)(double)((valPrice10 is DBNull) ? 0d : valPrice10)
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
                    var oldPriceDate = history.PriceValues[i - 1].Date;
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


        private void WriteAggregatedSheetAndCharts(IEnumerable<PriceHistory> histories)
        {
            var excelApp = new Application
            {
                Visible = true,
                WindowState = XlWindowState.xlMaximized,
                DisplayAlerts = false
            };
            var book = excelApp.Workbooks.Open(_excelFilePath);
            var aggregatedSheet = WriteAggregatedSheet(book, histories);
            var chartSheet = GenerateCharts(book, aggregatedSheet, histories);
            chartSheet.Activate();
        }

        private _Worksheet WriteAggregatedSheet(Workbook book, IEnumerable<PriceHistory> histories)
        {
            var sheets = book.Sheets;
            var hasAggregatedDataSheet = sheets.Count > 1;
            if(hasAggregatedDataSheet)
            {
                //Delete both aggData and charts sheets
                DeleteNextSheet();
                var hasChartSheet = sheets.Count > 1;
                if(hasChartSheet)
                {
                    DeleteNextSheet();
                }

                void DeleteNextSheet()
                {
                    _Worksheet toDelete = sheets[2];
                    toDelete.Delete();
                }
            }
            ((Worksheet)book.ActiveSheet).Columns.AutoFit();
            _Worksheet aggregatedSheet = sheets.Add(After: book.ActiveSheet);

            aggregatedSheet.Name = "AggData";
            aggregatedSheet.Cells[1, "A"].ColumnWidth = 1;
            aggregatedSheet.Cells[1, "A"].RowHeight = 10;

            for(int i = 0; i < histories.Count(); i++)
            {
                var history = histories.ElementAt(i);

                //Set item label
                Range titleRange = aggregatedSheet.Range[aggregatedSheet.Cells[2, 2 + _colPerTab * i], aggregatedSheet.Cells[2, _colPerTab + _colPerTab * i]];
                aggregatedSheet.Cells[2, 2 + _colPerTab * i] = history.Label;

                //Merge and center title cells
                titleRange.Merge();
                titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 16;

                //Set headers
                aggregatedSheet.Cells[3, 2 + _colPerTab * i] = "Date";
                aggregatedSheet.Cells[3, 3 + _colPerTab * i] = "x1";
                aggregatedSheet.Cells[3, 4 + _colPerTab * i] = "x10";
                aggregatedSheet.Cells[3, 5 + _colPerTab * i] = "x10/10";
                aggregatedSheet.Cells[3, 6 + _colPerTab * i] = $"Buy (Min + {buySellTresholdPercent}%)";
                aggregatedSheet.Cells[3, 7 + _colPerTab * i] = $"Sell (Max - {buySellTresholdPercent}%)";
                Range headersRange = aggregatedSheet.Range[aggregatedSheet.Cells[3, 2 + _colPerTab * i], aggregatedSheet.Cells[3, 7 + _colPerTab * i]];
                headersRange.Font.Size = 13;
                headersRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                //Write price values rows
                for(int j = 0; j < history.PriceValues.Count; j++)
                {
                    var price = history.PriceValues[j];

                    //Date
                    aggregatedSheet.Cells[4 + j, 2 + _colPerTab * i] = price.Date;

                    //x1
                    if(price.Price1 != 0)
                    {
                        aggregatedSheet.Cells[4 + j, 3 + _colPerTab * i] = price.Price1;
                    }

                    //x10 & x10/10
                    if(price.Price10 != 0)
                    {
                        aggregatedSheet.Cells[4 + j, 4 + _colPerTab * i] = price.Price10;
                        aggregatedSheet.Cells[4 + j, 5 + _colPerTab * i] = price.Price10On10;
                    }

                    //Min/Max
                    if(history.MinPrice != 0)
                    {
                        var minPrice = history.MinPrice;
                        var maxPrice = history.MaxPrice;
                        var diff = maxPrice - minPrice;


                        aggregatedSheet.Cells[4 + j, 6 + _colPerTab * i] = minPrice + (diff * buySellTresholdPercent / 100);
                        aggregatedSheet.Cells[4 + j, 7 + _colPerTab * i] = maxPrice - (diff * buySellTresholdPercent / 100);
                    }

                    //Spacing column
                    aggregatedSheet.Cells[4 + j, 8 + _colPerTab * i].ColumnWidth = 1;
                }

                //Set date format
                Range datesRange = aggregatedSheet.Range[aggregatedSheet.Cells[4, 2 + _colPerTab * i], aggregatedSheet.Cells[3 + history.PriceValues.Count, 2 + _colPerTab * i]];
                datesRange.NumberFormat = "DD/MM/YYYY";

                //Set borders
                Range wholeItemRange = aggregatedSheet.Range[aggregatedSheet.Cells[2, 2 + _colPerTab * i], aggregatedSheet.Cells[3 + history.PriceValues.Count, 7 + _colPerTab * i]];
                Borders borders1 = wholeItemRange.Borders;

                //All borders
                borders1.LineStyle = XlLineStyle.xlContinuous;
                borders1.Weight = 2d;

                //Thick borders around
                titleRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                headersRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                wholeItemRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            }

            aggregatedSheet.Columns.AutoFit();

            return aggregatedSheet;
        }

        private _Worksheet GenerateCharts(Workbook book, _Worksheet aggregatedSheet, IEnumerable<PriceHistory> histories)
        {
            var sheets = book.Sheets;
            _Worksheet chartsSheet = sheets.Add(After: aggregatedSheet);
            chartsSheet.Name = "Charts";

            var height = 275;
            var width = 685;
            object misValue = System.Reflection.Missing.Value;

            var alphabet = new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            for(int i = 0; i < histories.Count(); i++)
            {
                var history = histories.ElementAt(i);
                var top = 10 + 10 * (i / 2) + (i / 2) * height;
                var left = i % 2 == 0 ? 10 : 20 + width;

                //Create chart object
                ChartObjects xlCharts = (ChartObjects)chartsSheet.ChartObjects(Type.Missing);
                ChartObject myChart = xlCharts.Add(left, top, width, height);
                Chart chartPage = myChart.Chart;

                var colMult = _colPerTab * i;
                var dateAndPricesRange =
                    $"{GetAlphabetLetterForColumn(1 + colMult)}{3}" +
                    ":" +
                    $"{GetAlphabetLetterForColumn(2 + colMult)}{3 + history.PriceValues.Count}" +
                    ";" +
                    $"{GetAlphabetLetterForColumn(4 + colMult)}{3}" +
                    ":" +
                    $"{GetAlphabetLetterForColumn(6 + colMult)}{3 + history.PriceValues.Count}";

                string GetAlphabetLetterForColumn(int column)
                {
                    var currentLetter = alphabet[column % alphabet.Count()];

                    var baseLetterIdx = column / alphabet.Count();
                    if(baseLetterIdx != 0)
                    {
                        var baseLetter = alphabet[baseLetterIdx - 1];
                        currentLetter = baseLetter + currentLetter;
                    }

                    if(baseLetterIdx > alphabet.Count())
                    {
                        //Question isn't IF it's gonna break...
                        //Question is HOW it's gonna break...
                        //TODO: DEBUG & TEST
                    }
                    return currentLetter;
                }

                //Set chart data and style
                Range chartDataRange1 = aggregatedSheet.Range[dateAndPricesRange];
                chartPage.SetSourceData(chartDataRange1, misValue);
                chartPage.HasTitle = true;
                chartPage.ChartTitle.Text = history.Label;
                chartPage.ChartType = XlChartType.xlLine;
                chartPage.DisplayBlanksAs = XlDisplayBlanksAs.xlInterpolated;

                Series x1Serie = chartPage.FullSeriesCollection(1);
                x1Serie.Format.Line.ForeColor.RGB = (int)XlRgbColor.rgbPurple;
                x1Serie.Format.Line.Weight = 2f;

                Series x10On10Serie = chartPage.FullSeriesCollection(2);
                x10On10Serie.Format.Line.ForeColor.RGB = (int)XlRgbColor.rgbDarkTurquoise;
                x10On10Serie.Format.Line.Weight = 2f;

                Series minSerie = chartPage.FullSeriesCollection(3);
                minSerie.Name = "Buy";
                minSerie.Format.Line.ForeColor.RGB = (int)XlRgbColor.rgbGreen;
                minSerie.Format.Line.Weight = 1f;

                Series maxSerie = chartPage.FullSeriesCollection(4);
                maxSerie.Name = "Sell";
                maxSerie.Format.Line.ForeColor.RGB = (int)XlRgbColor.rgbRed;
                maxSerie.Format.Line.Weight = 1;


                var minValHistory = history.PriceValues.Where(x => x.Price1 != 0).Select(x => x.Price1)
                    .Union(history.PriceValues.Where(x => x.Price10On10 != 0).Select(x => x.Price10On10))
                    .Min();
                double minValHistoryRounded = minValHistory - (minValHistory % 1000);

                chartPage.Axes(XlAxisType.xlValue).MinimumScale = minValHistoryRounded;
            }

            return chartsSheet;
        }
    }
}
