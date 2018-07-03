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
            //string origin = "https://10.111.111.219:3000/";

            //EnableCorsAttribute cors = new EnableCorsAttribute(origin, "*", "GET,POST");

            config.EnableCors();
            // Web API configuration and services

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{*id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
