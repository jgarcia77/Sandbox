using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Sandbox.MVC.Startup))]
namespace Sandbox.MVC
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
