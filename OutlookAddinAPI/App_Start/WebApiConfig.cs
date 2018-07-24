using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Cors;

namespace OutlookAddinAPI
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            //string origin = "https://10.111.111.224:3000/";

            //EnableCorsAttribute cors = new EnableCorsAttribute("*", "*", "POST");

            config.EnableCors();
            // Web API configuration and services

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{action}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
