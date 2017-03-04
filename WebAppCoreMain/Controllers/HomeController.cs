using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;

namespace WebAppCoreMain.Controllers
{
    public class HomeController : Controller
    {
        private IMemoryCache _cache;

        public HomeController(IMemoryCache cache)
        {
            _cache = cache;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult About()
        {
            var key = "aboutMessage";

            var message = string.Empty;

            if (!_cache.TryGetValue(key, out message))
            {
                message = "Your application description page.";

                var options = new MemoryCacheEntryOptions().SetSlidingExpiration(TimeSpan.FromMinutes(1));

                _cache.Set(key, message, options);
            }

            ViewData["Message"] = message;

            return View();
        }

        public async Task<IActionResult> Contact()
        {
            var key = "contacttMessage";

            var message = await
                _cache.GetOrCreateAsync(key, entry => 
                {
                    entry.SlidingExpiration = TimeSpan.FromMinutes(1);

                    return Task.FromResult("Your contact page.");
                });

            ViewData["Message"] = message;

            return View();
        }

        public IActionResult Error()
        {
            return View();
        }
    }
}
