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

            //var properties = typeof(Instrument).GetProperties();

            //foreach (var property in properties)
            //{
            //    Console.WriteLine(property.Name);
            //}

            var currentUtc = DateTime.Now.ToUniversalTime();

            var currentTime = currentUtc.TimeOfDay;
            var currentHour = currentTime.Hours;

            DateTime startDateTime;

            if (currentHour < 4)
            {
                startDateTime = 
                    new DateTime(currentUtc.Year, currentUtc.Month, currentUtc.Day, 4, 0, 0);
            }
            else
            {
                var tomorrowUtc = currentUtc.AddDays(1);

                startDateTime = 
                    new DateTime(tomorrowUtc.Year, tomorrowUtc.Month, tomorrowUtc.Day, 4, 0, 0);
            }

            var dateTimeOffset = new DateTimeOffset(startDateTime);

            Console.WriteLine(currentTime);
            Console.WriteLine(currentHour);

            Console.Read();
        }
    }
}
