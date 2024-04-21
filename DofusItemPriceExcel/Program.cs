using DofusItemPriceExcel.Objects;
using Newtonsoft.Json;
using System;
using System.IO;

namespace DofusItemPriceExcelPj
{
    internal static class Program
    {
        private static readonly string AppdataDirectoryPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\DofusItemPrice";
        private static readonly string AppdataFileName = "lastPath.txt";
        private static string AppdataFilePath => AppdataDirectoryPath + AppdataFileName;

        static void Main(string[] args)
        {
            if(File.Exists(AppdataFilePath))
            {
                RunOptions options;
                try
                {
                    options = JsonConvert.DeserializeObject<RunOptions>(File.ReadAllText(AppdataFilePath));
                }
                catch(JsonReaderException e)
                {
                    options = new RunOptions();
                    if(e.Message.StartsWith("Unexpected character encountered while parsing value:"))
                    {
                        options.FilePath = File.ReadAllText(AppdataFilePath);
                    }
                }

                var runner = new ProgramRunner();
                runner.Run(options);
            }
        }
    }
}