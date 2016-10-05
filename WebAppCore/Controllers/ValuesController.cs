namespace WebAppCore.Controllers
{
    using System;
    using System.Collections.Generic;
    using Microsoft.AspNetCore.Mvc;
    using WebAppCore.Models;
    using WebAppCore.Attributes;

    [Route("api/[controller]")]
    //[ServiceFilter(typeof(LogExceptionAttribute))]
    public class ValuesController : Controller
    {
        // GET api/values
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        [HttpGet("{id}")]
        [ValidateModelState]
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        [HttpPost]
        [ValidateModelState]
        public IActionResult Post([FromBody]ValueModel model)
        {
            throw new NotImplementedException("Not implemented");

            return this.Ok(model);
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
