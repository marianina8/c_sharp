using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ITS.Startup))]
namespace ITS
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
