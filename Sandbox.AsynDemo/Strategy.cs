namespace Sandbox.AsynDemo
{
    using System;
    using System.Threading;

    public class Strategy
    {
        public void Execute(int callDuration, out int threadId)
        {
            Thread.Sleep(callDuration);

            Console.WriteLine("Execute method begins");
                        
            threadId = Thread.CurrentThread.ManagedThreadId;

            Console.WriteLine("Execute method ends");
        }

        public static void ExecuteStatic(int callDuration, out int threadId)
        {
            Thread.Sleep(callDuration);

            Console.WriteLine("Execute static method begins");

            threadId = Thread.CurrentThread.ManagedThreadId;

            Console.WriteLine("Execute static method ends");
        }

        public static Person GetPerson(int callDuration)
        {
            Thread.Sleep(callDuration);

            Console.WriteLine("Begin GetPerson");

            var person = new Person { Id = 1, Name = "Josue Garcia" };

            Console.WriteLine("End GetPerson");

            return person;
        }
    }
}
