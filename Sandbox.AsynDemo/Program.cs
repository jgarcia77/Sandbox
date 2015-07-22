using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox.AsynDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            ExecuteGetPerson();

            Console.Read();
        }

        static void ExecuteGetPerson()
        {
            Console.WriteLine("Begin ExecuteGetPerson");

            var caller = new AsynchronousGetPersonCaller(Strategy.GetPerson);

            Console.WriteLine("Begin Invoke");

            var result = caller.BeginInvoke(5000, new AsyncCallback(GetPersonCallback), null);

            Console.WriteLine("Passed Invoke");

            Console.WriteLine("End ExecuteGetPerson");
        }

        static void GetPersonCallback(IAsyncResult result)
        {
            Console.WriteLine("Begin GetPersonCallback");

            var asyncResult = (AsyncResult)result;

            var caller = (AsynchronousGetPersonCaller)asyncResult.AsyncDelegate;

            var person = caller.EndInvoke(result);

            Console.WriteLine("Id: {0}; Name: {1}", person.Id, person.Name);

            Console.WriteLine("End GetPersonCallback");
        }

        static void ExecuteStaticWithCallback()
        {
            Console.WriteLine("Begin ExecuteStaticWithOutWait");

            int threadId;

            var caller = new AsynchronousMethodCaller(Strategy.ExecuteStatic);

            Console.WriteLine("Begin Invoke");

            var result = caller.BeginInvoke(5000, out threadId, new AsyncCallback(CallbackMethod), null);

            Console.WriteLine("Passed Invoke");

            Console.WriteLine("End ExecuteStaticWithOutWait");
        }

        static void CallbackMethod(IAsyncResult ar)
        {
            Console.WriteLine("Begin CallbackMethod");

            Console.WriteLine("End CallbackMethod");
        }

        static void ExecuteInstanceMethod()
        {
            ExecuteWithWait();

            Console.WriteLine("------------------------------------------");

            ExecuteWithOutWait();
        }

        static void ExecuteStaticMethod()
        {
            
            ExecuteStaticWithOutWait();
        }

        static void ExecuteStaticWithOutWait()
        {
            Console.WriteLine("Begin ExecuteStaticWithOutWait");

            int threadId;

            var caller = new AsynchronousMethodCaller(Strategy.ExecuteStatic);

            Console.WriteLine("Begin Invoke");

            var result = caller.BeginInvoke(5000, out threadId, null, null);

            Console.WriteLine("Passed Invoke");

            Console.WriteLine("End ExecuteStaticWithOutWait");
        }

        static void ExecuteWithOutWait()
        {
            Console.WriteLine("Begin ExecuteWithOutWait");

            int threadId;

            var strategy = new Strategy();

            var caller = new AsynchronousMethodCaller(strategy.Execute);

            Console.WriteLine("Begin Invoke");

            var result = caller.BeginInvoke(5000, out threadId, null, null);

            Console.WriteLine("Passed Invoke");
            
            Console.WriteLine("End ExecuteWithOutWait");
        }

        static void ExecuteWithWait()
        {
            Console.WriteLine("Begin ExecuteWithWait");

            int threadId;

            var strategy = new Strategy();

            var caller = new AsynchronousMethodCaller(strategy.Execute);

            Console.WriteLine("Begin Invoke");

            var result = caller.BeginInvoke(5000, out threadId, null, null);

            Console.WriteLine("Passed Invoke");

            caller.EndInvoke(out threadId, result);

            Console.WriteLine("Thread Id: {0}", threadId);

            Console.WriteLine("End ExecuteWithWait");
        }

        public delegate void AsynchronousMethodCaller(int callDuration, out int threadId);
        public delegate Person AsynchronousGetPersonCaller(int callDuration);
    }
}
