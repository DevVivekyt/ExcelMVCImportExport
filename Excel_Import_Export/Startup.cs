using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Excel_Import_Export.Startup))]
namespace Excel_Import_Export
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
