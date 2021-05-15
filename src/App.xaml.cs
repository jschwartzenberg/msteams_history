using Microsoft.Identity.Client;
using System.Windows;

namespace MSTeamsHistory
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        static App()
        {
            _clientApp = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, Tenant)
                .Build();
            TokenCacheHelper.EnableSerialization(_clientApp.UserTokenCache);
        }

        // Below are the clientId (Application Id) of your app registration and the tenant information. 
        // You have to replace:
        // - the content of ClientID with the Application Id for your app registration
        // - Te content of Tenant by the information about the accounts allowed to sign-in in your application:
        //   - For Work or School account in your org, use your tenant ID, or domain
        //   - for any Work or School accounts, use organizations
        //   - for any Work or School accounts, or Microsoft personal account, use common
        //   - for Microsoft Personal account, use consumers
        private static string ClientId = "8f190bac-1f84-48e9-b981-efdea3a1a6b0";
        //private static string ClientId = "7c6d83df-d229-494b-8259-ed191b9fd479";
        // "0b8b0665-bc13-4fdc-bd72-e0227b9fc011";

        // Note: Tenant is important for the quickstart. We'd need to check with Andre/Portal if we
        // want to change to the AadAuthorityAudience.
        private static string Tenant = "common";
        //private static string Tenant = "luxoft.com";

        private static IPublicClientApplication _clientApp ;

        public static IPublicClientApplication PublicClientApp { get { return _clientApp; } }
    }
}
