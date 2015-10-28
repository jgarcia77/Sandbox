using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Sandbox.BackgroundWorkerDemo
{
    class Program
    {
        static BackgroundWorker bw;
        static string partNumber;
        static string description;

        static void Main(string[] args)
        {
            Console.WriteLine("Main start");

            Console.WriteLine("Enter part number: ");
            partNumber = Console.ReadLine();

            Console.WriteLine("Enter description: ");
            description = Console.ReadLine();

            bw = new BackgroundWorker();

            bw.DoWork += new DoWorkEventHandler(bw_DoWork);

            if (!bw.IsBusy)
            {
                bw.RunWorkerAsync();
            }

            Console.WriteLine("Main end");

            Console.Read();
        }

        static void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;

            searchProducts(worker);
        }

        static void searchProducts(BackgroundWorker worker)
        {
            Console.WriteLine("searchProducts start");

            Thread.Sleep(5000);

            Console.WriteLine("searchProducts end");
        }
    }
}
