using Microsoft.Identity.Client;

namespace AzureAdSSR.AccessToken
{
    public class TokenService
    {
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _tenantId;

        public TokenService(string clientId, string clientSecret, string tenantId)
        {
            _clientId = clientId;
            _clientSecret = clientSecret;
            _tenantId = tenantId;
        }

        public async Task<string> GetAccessToken(string refreshToken, string[] scopes)
        {
            var confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(_clientId)
                .WithClientSecret(_clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{_tenantId}"))
                .Build();

            var authenticationResult = await confidentialClientApplication.AcquireTokenSilent(scopes, refreshToken)
                .ExecuteAsync();

            return authenticationResult.AccessToken;
        }
    }
}
