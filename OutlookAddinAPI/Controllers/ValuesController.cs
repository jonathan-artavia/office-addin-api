using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace OutlookAddinAPI.Controllers
{
    public class ValuesController : ApiController
    {
        // GET api/values
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "No 'id' specified" };
        }
        // GET api/values/5
        [HttpGet()]
        [System.Web.Http.Cors.EnableCors("*", "*", "*")]
        public string Get(string id)
        {
            using (MailTrackerProvider prov = new MailTrackerProvider())
            {
                if (prov.OpenConnection())
                {
                    
                    return prov.IsTracked(System.Web.HttpUtility.UrlDecode(id)).ToString();
                }
                else
                {
                    return "false";
                }
            }
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody]string value)
        {
            
        }

        // PUT api/values/5
        [HttpPut]
        public void Put(string id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete()]
        public void Delete(int id)
        {
        }
    }
}
