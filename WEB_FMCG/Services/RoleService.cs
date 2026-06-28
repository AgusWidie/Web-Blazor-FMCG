using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.Auth;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface IRoleService
    {
        Task<PaginatedResult<Role>> GetPaged(int pageNumber, int pageSize);
        Task<List<Role>> GetAll();
        Task Create(Role a);
        Task Update(Role a);
        byte[] GenerateExportRole(List<Role> data);
    }

    public class RoleService : IRoleService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "master/api/v1/Role";
        public RoleService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<Role>> GetPaged(int pageNumber, int pageSize)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    PropertyNameCaseInsensitive = true
                };

                var token = await _auth.GetTokenAsync();
                if (string.IsNullOrWhiteSpace(token))
                {
                    throw new InvalidOperationException("User not logged in");
                }

                var url = $"{baseAPI}/GetPagedRole?pageNumber={pageNumber}&pageSize={pageSize}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<Role>>>(options);
                if (apiResponse == null || apiResponse.Data == null)
                    throw new InvalidOperationException("Data Null");
                return apiResponse.Data;
            }
            catch (HttpRequestException httpEx)
            {
                Console.WriteLine($"HTTP error: {httpEx.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General error: {ex.Message}");
                throw;
            }
        }


        public async Task<List<Role>> GetAll()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    PropertyNameCaseInsensitive = true
                };

                var token = await _auth.GetTokenAsync();
                if (string.IsNullOrWhiteSpace(token))
                {
                    throw new InvalidOperationException("User not logged in");
                }

                var request = new HttpRequestMessage(HttpMethod.Get, $"{baseAPI}/GetAllRole");
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);


                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<List<Role>>>(options);
                if (apiResponse == null || apiResponse.Data == null)
                    throw new InvalidOperationException("Data Null");
                return apiResponse.Data;

            }
            catch (HttpRequestException httpEx)
            {
                // Error dari HTTP, misal 401, 500, dll.
                Console.WriteLine($"HTTP error: {httpEx.Message}");
                throw; // optional: lempar lagi biar caller tahu
            }
            catch (Exception ex)
            {
                // Error lain, misal serialisasi JSON, jaringan, dll.
                Console.WriteLine($"General error: {ex.Message}");
                throw;
            }

        }

        public async Task Create(Role a)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    PropertyNameCaseInsensitive = true
                };

                var token = await _auth.GetTokenAsync();
                if (string.IsNullOrWhiteSpace(token))
                {
                    throw new InvalidOperationException("User not logged in");
                }

                var request = new HttpRequestMessage(HttpMethod.Post, $"{baseAPI}/CreateRole")
                {
                    Content = JsonContent.Create(a, options: options)
                };
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);

                // Lempar exception jika status code bukan 2xx
                response.EnsureSuccessStatusCode();
            }
            catch (HttpRequestException httpEx)
            {
                // Error dari HTTP, misal 401, 500, dll.
                Console.WriteLine($"HTTP error: {httpEx.Message}");
                throw; // optional: lempar lagi biar caller tahu
            }
            catch (Exception ex)
            {
                // Error lain, misal serialisasi JSON, jaringan, dll.
                Console.WriteLine($"General error: {ex.Message}");
                throw;
            }
        }

        public async Task Update(Role a)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    PropertyNameCaseInsensitive = true
                };

                var token = await _auth.GetTokenAsync();
                if (string.IsNullOrWhiteSpace(token))
                {
                    throw new InvalidOperationException("User not logged in");
                }

                var request = new HttpRequestMessage(HttpMethod.Put, $"{baseAPI}/UpdateRole")
                {
                    Content = JsonContent.Create(a)
                };
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);

                // Lempar exception jika status code bukan 2xx
                response.EnsureSuccessStatusCode();
            }
            catch (HttpRequestException httpEx)
            {
                // Error dari HTTP, misal 401, 500, dll.
                Console.WriteLine($"HTTP error: {httpEx.Message}");
                throw; // optional: lempar lagi biar caller tahu
            }
            catch (Exception ex)
            {
                // Error lain, misal serialisasi JSON, jaringan, dll.
                Console.WriteLine($"General error: {ex.Message}");
                throw;
            }
        }

        public byte[] GenerateExportRole(List<Role> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Role");

            // Header
            worksheet.Cells[1, 1].Value = "Role ID";
            worksheet.Cells[1, 2].Value = "Role Name";
            worksheet.Cells[1, 3].Value = "Description";
            worksheet.Cells[1, 4].Value = "Active";

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].RoleId;
                worksheet.Cells[i + 2, 2].Value = data[i].RoleName;
                worksheet.Cells[i + 2, 3].Value = data[i].Description;
                worksheet.Cells[i + 2, 4].Value = data[i].Active;
            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }
}
