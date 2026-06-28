using System.Net.Http.Headers;

namespace WEB_FMCG.Auth
{
    public class AuthMessageHandler : DelegatingHandler
    {
        private readonly AuthState _auth;

        public AuthMessageHandler(AuthState auth)
        {
            _auth = auth;
        }

        protected override async Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken)
        {
            if (_auth.IsAuthenticated)
            {
                request.Headers.Authorization =
                    new AuthenticationHeaderValue("Bearer", _auth.AccessToken);
            }

            return await base.SendAsync(request, cancellationToken);
        }
    }
}
