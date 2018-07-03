using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutlookAddinAPI.Models
{
    public class NotificationModel
    {
        public string SubscriptionId { get; set; }
        public string SubscriptionExpirationDateTime { get; set; }
        public int SequenceNumber { get; set; }
        public string ChangeType { get; set; }
        public string Resource { get; set; }
        public ResourceDataModel ResourceData { get; set; }
    }
}