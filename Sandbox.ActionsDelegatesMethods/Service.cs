namespace Sandbox.ActionsDelegatesMethods
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    public class Service
    {
        public async Task OutputMessage(string message)
        {
            await Task.Run(() => Console.WriteLine(message));
        }
    }
}
