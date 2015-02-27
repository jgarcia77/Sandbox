using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox.GuidGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var guid = Guid.NewGuid();

            Console.WriteLine(guid);

            guid = Guid.NewGuid();

            Console.WriteLine(guid);

            Console.Read();
        }
    }
}
