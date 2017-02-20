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
            NumberFormatting();
        }

        static void NumberFormatting()
        {
            while (true)
            {
                Console.Write("Enter a number: ");

                var input = Console.ReadLine();

                try
                {
                    var inputDouble = Convert.ToDouble(input);

                    var roundDouble = TruncateToSignificantDigits(inputDouble, 4);

                    Console.WriteLine(roundDouble);
                    Console.WriteLine(roundDouble.ToString("##,###.####"));
                }
                catch
                {
                    continue;
                }
            }
        }

        static double TruncateToSignificantDigits(double d, int digits)
        {
            var returnValue = 0.0;

            if (d != 0.0)
            {
                var absValue = Math.Abs(d);

                var log10Value = Math.Log10(absValue);

                var characteristicLength = Math.Floor(log10Value) + 1;

                var scale = Math.Pow(10, characteristicLength);

                var scaledValue = d / scale;

                var roundedScaledValue = Math.Round(scaledValue, digits, MidpointRounding.AwayFromZero);

                var rawValue = scale * roundedScaledValue;

                var mantissaLength = (int)characteristicLength >= digits ? 0 : digits - (int)characteristicLength;

                returnValue = Math.Round(rawValue, mantissaLength, MidpointRounding.AwayFromZero);
            }

            return returnValue;
        }

    }
}
