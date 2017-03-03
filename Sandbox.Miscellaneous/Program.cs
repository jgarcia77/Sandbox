using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Sandbox.Miscellaneous
{
    class Program
    {
        static void Main(string[] args)
        {
            //NumberFormatting();
            GetAnchorTags();

            Console.Read();
        }

        static void GetAnchorTags()
        {
            var value =
@"
The procedure could not be completed because the residuals demonstrate a significant pattern rather than random fluctuation. This can be caused by seasonal or other factors causing systematic changes in the relationship that are not included in the predicting variable. This pattern may be apparent from the plot of residuals. However, this model cannot be used for audit purposes.
 
Next Steps:
|Examine the plot of residuals to identify the possible source of the condition and make adjustments to the base data as needed.
|Consider introducing a new predicting variable (including a dummy variable if appropriate).
|For more information see examples provided in the Performing Substantive Analytical Procedures Guide section 3.4
|To discuss further alternatives, please contact support at <a href=""mailto: Revealsupport@deloitte.com"">Click to send email</a>.




If the sum of the residuals exceeds performance materiality, consider any of the other excesses that have been identified in the application. Reduce the cumulative residual by the amount of excesses that have been quantified and corroborated in step 1 below. If the remaining cumulative residual exceeds performance materiality, then seek additional explanations from management and quantify and corroborate the factors identified. If no specific items are identified, treat the remaining cumulative residual as a substantive analytical procedure misstatement.

Next Steps:
For each Excess, investigate the residual by:
|Inquiring of management and obtaining appropriate audit evidence relevant to management's response
If such investigation does not result in a satisfactory explanation, then the Excess is a substantive analytical procedure misstatement.
|Consider whether there are unusual patterns in the residuals that should be investigated (e.g., residuals tending strongly in one direction close to the individual thresholds, with a total that is multiples of performance materiality).
|Consider disaggregating the data, introducing one or more additional variables (real or dummy), or removing one or more variables that is causing the variability.
|For more information see examples provided in the Performing Substantive Analytical Procedures Guide section x.x Identification of Significant Differences. 
|To discuss further alternatives please submit an online support request <a href=""https://servicedesk.deloittenet.deloitte.com/CAisd/pdmweb.exe?OP=SHOW_DETAIL+PERSID=KD:448925+HTMPL=kt_document_view.htmpl+open_mode=2"" target=""_blank"">online support request</a>.
";

            //DumpHRefs(value);

            var results = Find(value);

            foreach (var item in results)
            {
                Console.WriteLine(item.Value);
                Console.WriteLine("HREF = {0}", item.Href);
                Console.WriteLine("TEXT = {0}", item.Text);
                Console.WriteLine(string.Empty);
            }
        }

        private static void DumpHRefs(string inputString)
        {
            Match m;
            string HRefPattern = "href\\s*=\\s*(?:[\"'](?<1>[^\"']*)[\"']|(?<1>\\S+))";

            try
            {
                m = Regex.Match(inputString, HRefPattern,
                                RegexOptions.IgnoreCase | RegexOptions.Compiled,
                                TimeSpan.FromSeconds(1));
                while (m.Success)
                {
                    Console.WriteLine("Found href " + m.Groups[1] + " at "
                       + m.Groups[1].Index);
                    m = m.NextMatch();
                }
            }
            catch (RegexMatchTimeoutException)
            {
                Console.WriteLine("The matching operation timed out.");
            }
        }

        public static List<LinkItem> Find(string file)
        {
            List<LinkItem> list = new List<LinkItem>();

            // 1.
            // Find all matches in file.
            MatchCollection m1 = Regex.Matches(file, @"(<a.*?>.*?</a>)",
                RegexOptions.Singleline);

            // 2.
            // Loop over each match.
            foreach (Match m in m1)
            {
                string value = m.Groups[1].Value;

                LinkItem i = new LinkItem { Value = value };

                // 3.
                // Get href attribute.
                Match m2 = Regex.Match(value, @"href=\""(.*?)\""",
                    RegexOptions.Singleline);
                if (m2.Success)
                {
                    i.Href = m2.Groups[1].Value;
                }

                // 4.
                // Remove inner tags from text.
                string t = Regex.Replace(value, @"\s*<.*?>\s*", "",
                    RegexOptions.Singleline);
                i.Text = t;

                list.Add(i);
            }
            return list;
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
            var returnValue = "0";

            if (value != 0)
            {
                var formatter = new StringBuilder();

                var valueArray = Math.Abs(value).ToString().Split('.');

                var characteristicLength = valueArray[0].Length;

                var mantissaLength = valueArray.Length == 1 ? 0 : valueArray[1].Length;

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
                
                var characteristicLength = absValue < Math.Pow(10, -6) ? 1 : Math.Floor(Math.Log10(absValue)) + 1;
                
                var scaledValue = value / Math.Pow(10, characteristicLength);

                var rawValue = Math.Floor(scaledValue * Math.Pow(10, digits) + 0.5);

                returnValue = rawValue / Math.Pow(10, digits - characteristicLength);
            }

            return returnValue;
        }

        static double RoundSignificantDigits_Original(double value, int digits)
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
