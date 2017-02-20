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

                    var roundDouble = RoundSignificantDigits(inputDouble, 4);

                    Console.WriteLine(roundDouble);
                    Console.WriteLine(FormatPrecise(roundDouble));
                }
                catch
                {
                    continue;
                }
            }
        }

        static string FormatPrecise(double value)
        {
            var returnValue = "0.0";

            if (value != 0.0)
            {
                var valueArray = Math.Abs(value).ToString().Split('.');

                var characteristicLength = 0;
                var mantissaLength = 0;
                
                if (valueArray.Length == 1)
                {
                    characteristicLength = valueArray[0].Length;
                    mantissaLength = 0;
                }
                else
                {
                    characteristicLength = valueArray[0].Length;
                    mantissaLength = valueArray[1].Length;
                }

                var formatter = new StringBuilder();

                var digitCounter = 0;

                for (var i = 0; i < characteristicLength; i++)
                {
                    digitCounter++;

                    if (digitCounter > 3)
                    {
                        digitCounter++;
                        formatter.Insert(0, ",");
                    }

                    formatter.Insert(0, "0");
                }

                if (mantissaLength != 0)
                {
                    formatter.Append(".");

                    for (var i = 0; i < mantissaLength; i++)
                    {
                        formatter.Append("0");
                    }
                }

                returnValue = value.ToString(formatter.ToString());
            }

            return returnValue;
        }

        static double RoundSignificantDigits(double value, int digits)
        {
            var returnValue = 0.0;

            if (value != 0.0)
            {
                var absValue = Math.Abs(value);

                var log10Value = Math.Log10(absValue);

                var characteristicLength = Math.Floor(log10Value) + 1;

                var scale = Math.Pow(10, characteristicLength);

                var scaledValue = value / scale;

                var roundedScaledValue = Math.Round(scaledValue, digits, MidpointRounding.AwayFromZero);

                var rawValue = scale * roundedScaledValue;

                var mantissaLength = (int)characteristicLength >= digits ? 0 : digits - (int)characteristicLength;

                returnValue = Math.Round(rawValue, mantissaLength, MidpointRounding.AwayFromZero);
            }

            return returnValue;
        }

    }
}
