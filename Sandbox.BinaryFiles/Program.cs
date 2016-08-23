using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Sandbox.BinaryFiles
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadBinaryFile2();

            Console.Read();
        }

        static void ReadBinaryFile()
        {
            var file = @"C:\Users\josueg\Documents\Projects\Marsh\TR000105.260";

            byte[] fileBytes = File.ReadAllBytes(file);
            StringBuilder sb = new StringBuilder();

            foreach (byte b in fileBytes)
            {
                var line = Convert.ToString(b, 2).PadLeft(8, '0');
                //sb.Append(line);
                Console.WriteLine(line);
            }

            //File.WriteAllText(outputFilename, sb.ToString());
        }

        static void ReadBinaryFile2()
        {
            var fs = new FileStream(@"C:\Users\josueg\Documents\Projects\Marsh\TR000105.260", FileMode.Open);
            var len = (int)fs.Length;
            var bits = new byte[len];
            fs.Read(bits, 0, len);
            // Dump 16 bytes per line
            for (int ix = 0; ix < len; ix += 16)
            {
                var cnt = Math.Min(16, len - ix);
                var line = new byte[cnt];
                Array.Copy(bits, ix, line, 0, cnt);
                // Write address + hex + ascii
                Console.Write("{0:X6}  ", ix);
                Console.Write(BitConverter.ToString(line));
                Console.Write("  ");
                // Convert non-ascii characters to .
                for (int jx = 0; jx < cnt; ++jx)
                    if (line[jx] < 0x20 || line[jx] > 0x7f) line[jx] = (byte)'.';
                Console.WriteLine(Encoding.ASCII.GetString(line));
            }
        }
    }
}
