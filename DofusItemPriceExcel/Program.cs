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
                var path = File.ReadAllText(AppdataFilePath);
                var runner = new ProgramRunner();
                runner.Run(path);
            }
        }
    }
}