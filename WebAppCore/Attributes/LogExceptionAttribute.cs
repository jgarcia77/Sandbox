namespace WebAppCore.Attributes
{
    using System;
    using Microsoft.AspNetCore.Mvc.Filters;
    using Microsoft.Extensions.Configuration;
    using Microsoft.AspNetCore.Mvc;
    using Models;

    public class LogExceptionAttribute : ExceptionFilterAttribute
    {
        private readonly IConfigurationRoot configuration;

        public LogExceptionAttribute(IConfigurationRoot configuration)
        {
            this.configuration = configuration;
        }

        public override void OnException(ExceptionContext context)
        {
            var model = new ResponseModel { IsSuccessful = false, Message = context.Exception.Message };
            var result = new OkObjectResult(model);
            context.Result = result;

            //var result = new ContentResult();
            //result.Content = context.Exception.Message;
            //context.Result = result;
        }
    }
}
