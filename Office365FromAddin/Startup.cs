using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Office365FromAddin.Startup))]
namespace Office365FromAddin
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            app.MapSignalR();
        }
    }
}
