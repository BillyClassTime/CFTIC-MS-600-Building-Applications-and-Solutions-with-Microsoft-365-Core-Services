using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Graph;
using System.Threading;


namespace Helpers
{
    public class AuthHandler : DelegatingHandler
    {
        private IAuthenticationProvider _authenticationProvider;

        public AuthHandler(IAuthenticationProvider authenticationProvider,
                           HttpMessageHandler innerHandler) {
            _authenticationProvider = authenticationProvider;
            InnerHandler = innerHandler;                    
        }
        protected override async Task<HttpResponseMessage> SendAsync (
            HttpRequestMessage requestMessage, CancellationToken cancellationToken)
        {
             await _authenticationProvider.AuthenticateRequestAsync(requestMessage);
             return await base.SendAsync(requestMessage,cancellationToken);
        }
    }
}