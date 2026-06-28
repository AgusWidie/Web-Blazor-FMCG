using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.Auth;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface IPhotoProductSecondaryDisplayService
    {
        Task<PaginatedResult<VPhotoProductSecondaryDisplay>> GetPaged(int pageNumber, int month, int year);
        Task<IEnumerable<VPhotoProductSecondaryDisplay>> PhotoProductSecondaryDisplayExport(int month, int year);
        Task Create(PhotoProductSecondaryDisplay a);
        byte[] GenerateExportPhotoProductSecondaryDisplay(List<VPhotoProductSecondaryDisplay> data);
    }

    public class PhotoProductSecondaryDisplayService : IPhotoProductSecondaryDisplayService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "transaction/api/v1/PhotoProductSecondaryDisplay";
        public PhotoProductSecondaryDisplayService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<VPhotoProductSecondaryDisplay>> GetPaged(int pageNumber, int month, int year)
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

                var url = $"{baseAPI}/GetPagedProductSecondaryDisplay?pageNumber={pageNumber}&month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<VPhotoProductSecondaryDisplay>>>(options);
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

        public async Task<IEnumerable<VPhotoProductSecondaryDisplay>> PhotoProductSecondaryDisplayExport(int month, int year)
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

                var url = $"{baseAPI}/GetProductSecondaryDisplayExport?month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<IEnumerable<VPhotoProductSecondaryDisplay>>>(options);
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

        public async Task Create(PhotoProductSecondaryDisplay a)
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

                var request = new HttpRequestMessage(HttpMethod.Post, $"{baseAPI}/CreateProductSecondaryDisplay")
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

        public byte[] GenerateExportPhotoProductSecondaryDisplay(List<VPhotoProductSecondaryDisplay> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("PhotoProductSecondaryDisplay");

            // Header
            worksheet.Cells[1, 1].Value = "Secondary Display ID";
            worksheet.Cells[1, 2].Value = "Sales ID";
            worksheet.Cells[1, 3].Value = "Sales Name";
            worksheet.Cells[1, 4].Value = "Store ID";
            worksheet.Cells[1, 5].Value = "Store Name";
            worksheet.Cells[1, 6].Value = "Url Photo";
            worksheet.Cells[1, 7].Value = "Photo Date";
            worksheet.Cells[1, 8].Value = "Month";
            worksheet.Cells[1, 9].Value = "Year";


            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].SecondaryDisplayId;
                worksheet.Cells[i + 2, 2].Value = data[i].SalesId;
                worksheet.Cells[i + 2, 3].Value = data[i].SalesName;
                worksheet.Cells[i + 2, 4].Value = data[i].StoreId;
                worksheet.Cells[i + 2, 5].Value = data[i].StoreName;
                worksheet.Cells[i + 2, 6].Value = data[i].UrlPhoto;
                worksheet.Cells[i + 2, 7].Value = data[i].PhotoDate != null ? Convert.ToDateTime(data[i].PhotoDate).ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cells[i + 2, 8].Value = data[i].Month;
                worksheet.Cells[i + 2, 9].Value = data[i].Year;

            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }
}
