using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutlookAddinAPI.Models
{
    public class ResponseModel<T>
    {
        public List<T> Value { get; set; }
    }
}