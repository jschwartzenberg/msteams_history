using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace MSTeamsHistory
{

    public class TokenHolder
    {
        private AuthenticationResult authResult = null;
        private IPublicClientApplication app;
        private readonly string[] scopes;
        private readonly IAccount account;

        public TokenHolder(IPublicClientApplication app, string[] scopes, AuthenticationResult authResult)
        {
            this.app = app;
            this.scopes = scopes;
            this.account = authResult.Account;
            this.authResult = authResult;
        }

        public string getToken()
        {
            lock (authResult)
            {
                if (DateTimeOffset.Compare(DateTimeOffset.UtcNow, authResult.ExpiresOn.AddMinutes(-5)) < 0)
                {
                    return authResult.AccessToken;
                }
                else
                {
                    authResult = app.AcquireTokenSilent(scopes, account).ExecuteAsync().Result;
                    return authResult.AccessToken;
                }
            }
        }

        public void refreshToken()
        {
            lock (authResult)
            {
                authResult = app.AcquireTokenSilent(scopes, account).ExecuteAsync().Result;
            }
        }
    }
}
