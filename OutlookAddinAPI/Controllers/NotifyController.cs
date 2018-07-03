﻿using Newtonsoft.Json;
using OutlookAddinAPI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

namespace OutlookAddinAPI.Controllers
{
    public class NotifyController : ApiController
    {
        public const string NOTIF_ID_GUID = "{B9EDF33A-D5BE-4B2E-8D0B-4B52EA010E92}";
        /// <summary>
        /// Responds to requests generated by subscriptions registered with
        /// the Outlook Notifications REST API. 
        /// </summary>
        /// <param name="validationToken">The validation token sent by Outlook when
        /// validating the Notification URL for the subscription.</param>
        public async Task<HttpResponseMessage> Post(string validationToken = null)
        {
            // If a validation token is present, we need to respond within 5 seconds.
            if (validationToken != null)
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);
                response.Content = new StringContent(validationToken);
                return response;
            }

            // Present only if the client specified the ClientState property in the 
            // subscription request. 
            IEnumerable<string> clientStateValues;
            Request.Headers.TryGetValues("ClientState", out clientStateValues);

            if (clientStateValues != null)
            {
                var clientState = clientStateValues.ToList().FirstOrDefault();
                if (clientState != null)
                {
                    // TODO: Use the client state to verify the legitimacy of the notification.
                }
            }

            // Read and parse the request body.
            var content = await Request.Content.ReadAsStringAsync();
            var notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(content);

            // TODO: Do something with the notification.

            return new HttpResponseMessage(HttpStatusCode.OK);
        }
    }
}
