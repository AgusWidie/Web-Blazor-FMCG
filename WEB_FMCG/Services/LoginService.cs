using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.Auth;
using WEB_FMCG.DTO;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface ILoginService
    {
        Task<ApiLoginResponse> LoginUser(LoginRequest a);
    }

    public class LoginService : ILoginService
    {

        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "register/api/v1";
        public LoginService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<ApiLoginResponse> LoginUser(LoginRequest a)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    PropertyNameCaseInsensitive = true
                };

                var request = new HttpRequestMessage(HttpMethod.Post, $"{baseAPI}/Login/LoginUser")
                {
                    Content = JsonContent.Create(a, options: options)
                };

                var response = await _http.SendAsync(request);

                // Baca response body dulu
                var responseBody = await response.Content.ReadAsStringAsync();
                //Console.WriteLine("StatusCode: " + response.StatusCode);
                //Console.WriteLine("Response Body: " + responseBody);

                // Jika status bukan 2xx, lempar exception dengan body
                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"API Error: {response.StatusCode} - {responseBody}");
                }

                // Deserialize
                var loginResponse = await response.Content.ReadFromJsonAsync<ApiLoginResponse>(options);
                await _auth.SetTokenAsync(loginResponse.Data.SalesId, loginResponse.Data.UserName, loginResponse.Data.Token, loginResponse.Data.RoleId);
                
                return loginResponse;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Login error: {ex.Message}");
                throw;
            }
        }

    }
}
