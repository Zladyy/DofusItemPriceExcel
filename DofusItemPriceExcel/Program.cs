using System;

namespace DofusItemPriceExcelPj
{
    internal static class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("App starting");

            var runner = new ProgramRunner();
            runner.Run();

            Console.ReadKey();
        }
    }
}