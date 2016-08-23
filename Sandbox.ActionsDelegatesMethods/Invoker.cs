namespace Sandbox.ActionsDelegatesMethods
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Text;
    using System.Threading.Tasks;

    public class Invoker
    {
        public static Guid Execute(Expression<Action> method)
        {
            var returnValue = Guid.Empty;

            method.Compile().Invoke();

            returnValue = Guid.NewGuid();

            return returnValue;
        }

        public static Guid ExecuteAction(Action method)
        {
            Console.WriteLine("Execute Action Now");
            return Execute(() => method.BeginInvoke(null, null));
        }
    }
}
