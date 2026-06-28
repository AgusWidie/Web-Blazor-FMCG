using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.DTO;

namespace WEB_FMCG.Services
{
    public interface ILogoutService
    {
        Task<ApiLoginResponse> LogoutUser(LoginRequest a);
    }

    public class LogoutService : ILogoutService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "register/api/v1";
        public LogoutService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<ApiLoginResponse> LogoutUser(LoginRequest a)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    PropertyNameCaseInsensitive = true
                };

                var request = new HttpRequestMessage(HttpMethod.Post, $"{baseAPI}/Logout/LogoutUser")
                {
                    Content = JsonContent.Create(a, options: options)
                };

                var response = await _http.SendAsync(request);
                var responseBody = await response.Content.ReadAsStringAsync();
                //Console.WriteLine("StatusCode: " + response.StatusCode);
                //Console.WriteLine("Response Body: " + responseBody);

                // Jika status bukan 2xx, lempar exception dengan body
                if (!response.IsSuccessStatusCode) { 
                    throw new Exception($"API Error: {response.StatusCode} - {responseBody}");
                }

                // Deserialize
                var loginResponse = await response.Content.ReadFromJsonAsync<ApiLoginResponse>(options);
#pragma warning disable CS8603 // Possible null reference return.
                return loginResponse;
#pragma warning restore CS8603 // Possible null reference return.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"LogoutUser error: {ex.Message}");
                throw;
            }
        }
    }
}
