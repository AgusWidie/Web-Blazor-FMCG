using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.Auth;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface IAreaService
    {
        Task<PaginatedResult<VArea>> GetPaged(int pageNumber, int pageSize);
        Task<List<VArea>> GetAll();
        Task Create(Area a);
        Task Update(Area a);
        byte[] GenerateExportArea(List<VArea> data);
    }

    public class AreaService : IAreaService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "master/api/v1/Area";

        public AreaService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<VArea>> GetPaged(int pageNumber, int pageSize)
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

                var url = $"{baseAPI}/GetPagedArea?pageNumber={pageNumber}&pageSize={pageSize}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                // Asumsikan API mengembalikan object PaginatedResult<VArea>
                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<VArea>>>(options);
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


        public async Task<List<VArea>> GetAll()
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

                var request = new HttpRequestMessage(HttpMethod.Get, $"{baseAPI}/GetAllArea");
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<List<VArea>>>(options);
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

        public async Task Create(Area a)
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

                var request = new HttpRequestMessage(HttpMethod.Post, $"{baseAPI}/CreateArea")
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

        public async Task Update(Area a)
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

                var request = new HttpRequestMessage(HttpMethod.Put, $"{baseAPI}/UpdateArea")
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

        public byte[] GenerateExportArea(List<VArea> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Area");

            // Header
            worksheet.Cells[1, 1].Value = "Region ID";
            worksheet.Cells[1, 2].Value = "Region Name";
            worksheet.Cells[1, 3].Value = "Area ID";
            worksheet.Cells[1, 4].Value = "Area Name";
            worksheet.Cells[1, 5].Value = "Active";
            worksheet.Cells[1, 6].Value = "Created Date";
            worksheet.Cells[1, 7].Value = "Created By";
            worksheet.Cells[1, 8].Value = "Updated Date";
            worksheet.Cells[1, 9].Value = "Updated By";

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].RegionId;
                worksheet.Cells[i + 2, 2].Value = data[i].RegionName;
                worksheet.Cells[i + 2, 3].Value = data[i].AreaId;
                worksheet.Cells[i + 2, 4].Value = data[i].AreaName;
                worksheet.Cells[i + 2, 5].Value = data[i].Active;
                worksheet.Cells[i + 2, 6].Value = data[i].CreatedDate != null ? Convert.ToDateTime(data[i].CreatedDate).ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cells[i + 2, 7].Value = data[i].CreatedBy;
                worksheet.Cells[i + 2, 8].Value = data[i].UpdatedDate != null ? Convert.ToDateTime(data[i].UpdatedDate).ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cells[i + 2, 9].Value = data[i].UpdatedBy;
            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }
}
