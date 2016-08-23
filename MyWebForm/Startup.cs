using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(MyWebForm.Startup))]
namespace MyWebForm
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
