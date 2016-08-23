using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox.ActionsDelegatesMethods
{
    class Program
    {
        static void Main(string[] args)
        {
            //TestExpressionAction();

            Console.Read();
        }

        //static async Task TestExpressionAction()
        //{ 
        //    var service = new Service();

        //    var result1 = Invoker.Execute(() => { await service.OutputMessage("Test Expression Action 1"); });
        //}

        static void TestAction()
        {
            var service = new Service();
            //var action = new Action(service.OutputMessage("test"));
            //var result2 = Invoker.ExecuteAction(action);
        }
    }
}
