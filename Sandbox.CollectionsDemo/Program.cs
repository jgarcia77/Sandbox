using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox.CollectionsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var collectoin = GetInts();

            Console.WriteLine(collectoin.Count());

            foreach (var i in collectoin)
            {
                Console.WriteLine(i);

                Console.WriteLine(collectoin.Count());
            }

            Console.Read();
        }

        static IEnumerable<int> GetInts()
        {
            var collection = new List<int>();

            for (var i = 1; i <= 10; i++)
            {
                collection.Add(i);
            }

            return collection.AsEnumerable();
        }
    }
}
