using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using GetCountersLib;
using ReactiveUI;

namespace TestConsoleApp
{
    static class Program
    {
        private static async Task Main(string[] args)
        {
            Console.WriteLine($"start {DateTime.Now}");
            var start = new Starter(40, 180, 5);

            Task.Run(() => start.AddCountersAsync());

            Console.WriteLine("ob");
            var ob = start.CreateObservable();

            Console.ReadLine();
            Console.WriteLine("read line");
            ob.Dispose();
            start.CloseExcelFile();
            Console.ReadLine();
        }
    }
}
