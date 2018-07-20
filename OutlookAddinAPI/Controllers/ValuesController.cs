using OutlookAddinAPI.Models;
using System;
using System.Collections.Generic;
using System.IO;
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
        // GET api/values/tracked?id=5
        [HttpPost]
        [System.Web.Http.Cors.EnableCors("*", "*", "*"), ActionName("tracked")]
        public string Get(string id)
        {
            string body = string.Empty;
            Stream s = this.Request.Content.ReadAsStreamAsync().Result;
            s.Seek(0, SeekOrigin.Begin);
            using (var stream = new StreamReader(s))
            {
                //stream.Seek(0, SeekOrigin.Begin);
                body = stream.ReadToEnd(); //Newtonsoft.Json.JsonConvert.SerializeObject(re);;
            }
            
            using (MailTrackerProvider prov = new MailTrackerProvider())
            {
                if (!string.IsNullOrEmpty(body))
                {
                    if (prov.OpenConnection())
                    {

                        if (prov.IsTracked(System.Web.HttpUtility.UrlDecode(id)))
                        {
                            System.Data.DataTable tb = prov.FindByConversationId(id);
                            prov.SaveEmail(body, long.Parse(tb.Rows[0]["ID"].ToString()));
                            return Newtonsoft.Json.JsonConvert.SerializeObject(new { IsTracked = true } );
                        }
                        return Newtonsoft.Json.JsonConvert.SerializeObject(new { IsTracked = false });
                    }
                    else
                    {
                        return Newtonsoft.Json.JsonConvert.SerializeObject(new { IsTracked = false });
                    }
                }
                else
                {
                    throw new InvalidCastException("Mail JSON string arrived empty or with wrong format");
                }
            }
        }

        [HttpPost]
        [ActionName("stracker")]
        public string StartStopTracker(bool start)
        {
            string body = string.Empty;
            Stream s = this.Request.Content.ReadAsStreamAsync().Result;
            s.Seek(0, SeekOrigin.Begin);
            using (var stream = new StreamReader(s))
            {
                //stream.Seek(0, SeekOrigin.Begin);
                body = stream.ReadToEnd(); //Newtonsoft.Json.JsonConvert.SerializeObject(re);;
            }
            try
            {

                if (!string.IsNullOrEmpty(body))
                {
                    MailItem usItem = Newtonsoft.Json.JsonConvert.DeserializeObject<MailItem>(body); //MailItem.FromJson(body);
                    if (usItem != null)
                    {
                        using (MailTrackerProvider prov = new MailTrackerProvider())
                        {
                            prov.OpenConnection();
                            if (start)
                            {
                                if (prov.StartTracker(body, usItem.DataP0.DataP0.ConversationId, usItem.DisplayName) > 0)
                                {
                                    return "Started";
                                }
                            }
                            else
                            {
                                if (prov.StopTracker(body, usItem.DataP0.DataP0.ConversationId, usItem.DisplayName) > 0)
                                {
                                    return "Stopped";
                                }
                            }
                        }
                    }
                    else
                    {
                        throw new InvalidCastException("Error: Mail JSON string arrived with wrong format");
                    }
                }
                else
                {
                    throw new InvalidCastException("Error: Mail JSON string arrived empty or with wrong format");
                }
                return "no.";
            }
            catch (Exception ex)
            {
                //MailItem usItem = Newtonsoft.Json.JsonConvert.DeserializeObject<MailItem>(body);
                //return "Error: " + ex.Message + "\r\n\r\n" + Newtonsoft.Json.JsonConvert.SerializeObject(usItem) + "\r\n\r\n" + ex.StackTrace;
                throw ex;
            }
        }
    }
}
