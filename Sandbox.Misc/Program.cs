using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace Sandbox.Misc
{
    class Program
    {
        static void Main(string[] args)
        {
            //var properties = typeof(Ticket).GetProperties();

            //foreach (var property in properties)
            //{
            //    Console.WriteLine(property.Name);
            //}

            var properties = typeof(Instrument).GetProperties();

            foreach (var property in properties)
            {
                Console.WriteLine(property.Name);
            }

            Console.Read();
        }
    }
}
