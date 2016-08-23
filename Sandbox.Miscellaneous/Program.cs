using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox.Miscellaneous
{
    class Program
    {
        static void Main(string[] args)
        {
            int dataTypeId = 2;

            var dataType = (DataTypes)dataTypeId;

            var dataTypeId2 = (int)dataType;
        }
    }

    enum DataTypes { DateTime, String, Boolean }
}
