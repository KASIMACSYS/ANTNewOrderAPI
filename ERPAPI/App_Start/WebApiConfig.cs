using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Cors;

namespace ERPAPI
{
    public static class WebApiConfig
    {
        private static string GetAllowedOrigins()
        {
            //Make a call to the database to get allowed origins and convert to a comma separated string
            //return "http://www.example.com,http://localhost:59452,http://localhost:25495";
            return "http://localhost,http://192.168.1.2,http://localhost:4200";
        }

        public static void Register(HttpConfiguration config)
        {
            string origins = GetAllowedOrigins();
            var cors = new EnableCorsAttribute(origins, "*", "*");
            //var cors = new EnableCorsAttribute("http://localhost:4200", "*", "*");
            config.EnableCors(cors);

            // Web API configuration and services

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
