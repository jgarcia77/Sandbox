using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace Sandbox.WebApiJScroll.Controllers
{
    public class ValuesController : ApiController
    {
        [HttpGet]
        [Route("api/numbers/{start}/{end}")]
        public IHttpActionResult GetNumbers(int start, int end)
        {
            var returnList = new List<int>();

            for (var i = start; i <= end; i++)
            {
                returnList.Add(i);
            }

            return Ok(returnList);
        }

        // GET api/values
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        public void Post([FromBody]string value)
        {
        }

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}
