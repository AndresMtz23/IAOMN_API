using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;

namespace IAOMN_API_REST.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    [RoutePrefix("v1/Proyecto/Northwind")]
    public class NorthwindController : ApiController
    {
        [HttpGet]
        [Route("Testing")]
        public HttpResponseMessage Testing()
        {
            return Request.CreateResponse(HttpStatusCode.OK, "Funciona!");
        }
    }
}
