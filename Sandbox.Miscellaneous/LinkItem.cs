using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox.Miscellaneous
{
    public struct LinkItem
    {
        public string Value;
        public string Href;
        public string Text;

        public override string ToString()
        {
            return Href + "\n\t" + Text;
        }
    }
}
