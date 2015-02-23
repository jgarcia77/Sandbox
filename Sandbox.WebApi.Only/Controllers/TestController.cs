using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;

namespace Sandbox.WebApi.Only.Controllers
{
    [EnableCors(origins: "http://localhost:56195", headers: "origin,Access-Control-Allow-Origin", methods: "*")]
    public class TestController : ApiController
    {
        
        public string Get()
        {
            return "This is the test controller";
        }
    }
}
