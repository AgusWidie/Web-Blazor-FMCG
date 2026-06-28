using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.Auth;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface IPhotoProductPriceCompetitorService
    {
        Task<PaginatedResult<VPhotoProductPriceCompetitor>> GetPaged(int pageNumber, int month, int year);
        Task<IEnumerable<VPhotoProductPriceCompetitor>> PhotoProductPriceCompetitorExport(int month, int year);
        Task Create(PhotoProductPriceCompetitor a);
        byte[] GenerateExportPhotoProductPriceCompetitor(List<VPhotoProductPriceCompetitor> data);
    }

    public class PhotoProductPriceCompetitorService : IPhotoProductPriceCompetitorService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "transaction/api/v1/PhotoProductPriceCompetitor";
        public PhotoProductPriceCompetitorService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<VPhotoProductPriceCompetitor>> GetPaged(int pageNumber, int month, int year)
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

                var url = $"{baseAPI}/GetPagedProductPriceCompetitor?pageNumber={pageNumber}&month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<VPhotoProductPriceCompetitor>>>(options);
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

        public async Task<IEnumerable<VPhotoProductPriceCompetitor>> PhotoProductPriceCompetitorExport(int month, int year)
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

                var url = $"{baseAPI}/GetProductPriceCompetitorExport?month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<IEnumerable<VPhotoProductPriceCompetitor>>>(options);
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

        public async Task Create(PhotoProductPriceCompetitor a)
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

                var request = new HttpRequestMessage(HttpMethod.Post, $"{baseAPI}/CreateProductPriceCompetitor")
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

        public byte[] GenerateExportPhotoProductPriceCompetitor(List<VPhotoProductPriceCompetitor> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("PhotoProductPriceCompetitor");

            // Header
            worksheet.Cells[1, 1].Value = "Competitor ID";
            worksheet.Cells[1, 2].Value = "Sales ID";
            worksheet.Cells[1, 3].Value = "Sales Name";
            worksheet.Cells[1, 4].Value = "Store ID";
            worksheet.Cells[1, 5].Value = "Store Name";
            worksheet.Cells[1, 6].Value = "Product ID";
            worksheet.Cells[1, 7].Value = "Product Name";
            worksheet.Cells[1, 8].Value = "Price Competitor";
            worksheet.Cells[1, 9].Value = "Url Photo";
            worksheet.Cells[1, 10].Value = "Month";
            worksheet.Cells[1, 11].Value = "Year";


            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].CompetitorId;
                worksheet.Cells[i + 2, 2].Value = data[i].SalesId;
                worksheet.Cells[i + 2, 3].Value = data[i].SalesName;
                worksheet.Cells[i + 2, 4].Value = data[i].StoreId;
                worksheet.Cells[i + 2, 5].Value = data[i].StoreName;
                worksheet.Cells[i + 2, 6].Value = data[i].ProductId;
                worksheet.Cells[i + 2, 7].Value = data[i].ProductName;
                worksheet.Cells[i + 2, 8].Value = data[i].PriceCompetitor;
                worksheet.Cells[i + 2, 9].Value = data[i].UrlPhoto;
                worksheet.Cells[i + 2, 10].Value = data[i].Month;
                worksheet.Cells[i + 2, 11].Value = data[i].Year;

            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }

}

