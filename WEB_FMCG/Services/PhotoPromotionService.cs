using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.Auth;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface IPhotoPromotionService
    {
        Task<PaginatedResult<VPhotoPromotion>> GetPaged(int pageNumber, int month, int year);
        Task<IEnumerable<VPhotoPromotion>> PhotoPromotionExport(int month, int year);
        Task Create(PhotoPromotion a);
        byte[] GenerateExportPhotoPromotion(List<VPhotoPromotion> data);
    }

    public class PhotoPromotionService : IPhotoPromotionService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "transaction/api/v1/PhotoPromotion";
        public PhotoPromotionService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<VPhotoPromotion>> GetPaged(int pageNumber, int month, int year)
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

                var url = $"{baseAPI}/GetPagedPromotion?pageNumber={pageNumber}&month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<VPhotoPromotion>>>(options);
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

        public async Task<IEnumerable<VPhotoPromotion>> PhotoPromotionExport(int month, int year)
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

                var url = $"{baseAPI}/GetPromotionExport?month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<IEnumerable<VPhotoPromotion>>>(options);
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

        public async Task Create(PhotoPromotion a)
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

                var request = new HttpRequestMessage(HttpMethod.Post, $"{baseAPI}/CreatePromotion")
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

        public byte[] GenerateExportPhotoPromotion(List<VPhotoPromotion> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("PhotoProductPromotion");

            // Header
            worksheet.Cells[1, 1].Value = "Promotion ID";
            worksheet.Cells[1, 2].Value = "Sales ID";
            worksheet.Cells[1, 3].Value = "Sales Name";
            worksheet.Cells[1, 4].Value = "Store ID";
            worksheet.Cells[1, 5].Value = "Store Name";
            worksheet.Cells[1, 6].Value = "Promotion Name";
            worksheet.Cells[1, 7].Value = "Weekly Promotion Fees";
            worksheet.Cells[1, 8].Value = "Budget Promotion Store";
            worksheet.Cells[1, 9].Value = "Url Photo";
            worksheet.Cells[1, 10].Value = "Start Date";
            worksheet.Cells[1, 11].Value = "End Date";
            worksheet.Cells[1, 12].Value = "Active";
            worksheet.Cells[1, 13].Value = "Description";
            worksheet.Cells[1, 14].Value = "Month";
            worksheet.Cells[1, 15].Value = "Year";


            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].PromotionId;
                worksheet.Cells[i + 2, 2].Value = data[i].SalesId;
                worksheet.Cells[i + 2, 3].Value = data[i].SalesName;
                worksheet.Cells[i + 2, 4].Value = data[i].StoreId;
                worksheet.Cells[i + 2, 5].Value = data[i].StoreName;
                worksheet.Cells[i + 2, 6].Value = data[i].PromotionName;
                worksheet.Cells[i + 2, 7].Value = data[i].WeeklyPromotionFees;
                worksheet.Cells[i + 2, 8].Value = data[i].BudgetPromotionStore;
                worksheet.Cells[i + 2, 9].Value = data[i].UrlPhoto;
                worksheet.Cells[i + 2, 10].Value = data[i].StartDate != null ? Convert.ToDateTime(data[i].StartDate).ToString("yyyy-MM-dd") : "";
                worksheet.Cells[i + 2, 11].Value = data[i].EndDate != null ? Convert.ToDateTime(data[i].EndDate).ToString("yyyy-MM-dd") : "";
                worksheet.Cells[i + 2, 12].Value = data[i].Active;
                worksheet.Cells[i + 2, 13].Value = data[i].Description;
                worksheet.Cells[i + 2, 14].Value = data[i].Month;
                worksheet.Cells[i + 2, 15].Value = data[i].Year;

            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }
}
