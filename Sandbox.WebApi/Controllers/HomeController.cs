using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using Sandbox.WebApi.Models;
using System.Configuration;

namespace Sandbox.WebApi.Controllers
{
    public class HomeController : Controller
    {
        private const string C_ANGULARAPP = "myApp.settings";
        private const string C_ANGULARJS_WEBAPIBASEURL = "AngularJS.WebApiBaseUrl";


        public ActionResult Index()
        {
            var settings = new Settings
            {
                WebApiBaseUrl = GetAppSetting<string>(C_ANGULARJS_WEBAPIBASEURL)
            };

            var serializerSettings = new JsonSerializerSettings
            {
                ContractResolver = new CamelCasePropertyNamesContractResolver()
            };
            var settingsJson = JsonConvert.SerializeObject(settings, Formatting.Indented, serializerSettings);

            var settingsVm = new SettingsViewModel
            {
                SettingsJson = settingsJson,
                AngularModuleName = C_ANGULARAPP
            };

            ViewBag.Title = "Home Page";

            return View(settingsVm);
        }

        protected static T GetAppSetting<T>(string key)
        {
            return (T)Convert.ChangeType(ConfigurationManager.AppSettings[key], typeof(T));
        }
    }
}
