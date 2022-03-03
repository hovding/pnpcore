using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PnP.Core.Auth;
using System.Security.Cryptography.X509Certificates;

[assembly: FunctionsStartup(typeof(Demo.AzureFunction.Startup))]

namespace Demo.AzureFunction
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            var config = builder.GetContext().Configuration;
            var azureFunctionSettings = new AzureFunctionSettings();
            azureFunctionSettings.TenantId = "769c8c85-ac79-4d8b-b09c-e5d139b21580";
            azureFunctionSettings.ClientId = "597f8a5b-f0a9-43fd-8727-d0ee3cbebf5e";
            azureFunctionSettings.CertificateThumbprint = "B904DE61200496BAAB2B9A4F066873494EDFAC1C";
            azureFunctionSettings.SiteUrl = "https://m5b3.sharepoint.com/sites/Dokstyring";
            config.Bind(azureFunctionSettings);

            builder.Services.AddPnPCore(options =>
            {
                // Disable telemetry because of mixed versions on AppInsights dependencies
                options.DisableTelemetry = true;

                // Configure an authentication provider with certificate (Required for app only)
                var authProvider = new X509CertificateAuthenticationProvider(azureFunctionSettings.ClientId,
                    azureFunctionSettings.TenantId,
                    StoreName.My,
                    StoreLocation.CurrentUser,
                    azureFunctionSettings.CertificateThumbprint);
                // And set it as default
                options.DefaultAuthenticationProvider = authProvider;

                // Add a default configuration with the site configured in app settings
                options.Sites.Add("Default",
                       new PnP.Core.Services.Builder.Configuration.PnPCoreSiteOptions
                       {
                           SiteUrl = azureFunctionSettings.SiteUrl,
                           AuthenticationProvider = authProvider
                       });
            });
        }
    }
}