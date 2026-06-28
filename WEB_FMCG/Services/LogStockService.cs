using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.Auth;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface ILogStockService
    {
        Task<PaginatedResult<VLogStock>> GetPaged(int pageNumber, int pageSize, int month, int year);
        Task<IEnumerable<VLogStock>> LogStockExport(int month, int year);
        byte[] GenerateExportLogStock(List<VLogStock> data);
    }

    public class LogStockService : ILogStockService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "transaction/api/v1/LogStock";
        public LogStockService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<VLogStock>> GetPaged(int pageNumber, int pageSize, int month, int year)
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

                var url = $"{baseAPI}/GetPagedLogStock?pageNumber={pageNumber}&pageSize={pageSize}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<VLogStock>>>(options);
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

        public async Task<IEnumerable<VLogStock>> LogStockExport(int month, int year)
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

                var url = $"{baseAPI}/GetLogStockExport?month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<IEnumerable<VLogStock>>>(options);
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

        public byte[] GenerateExportLogStock(List<VLogStock> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("LogStock");

            // Header
            worksheet.Cells[1, 1].Value = "Product Name";
            worksheet.Cells[1, 2].Value = "Store Name";
            worksheet.Cells[1, 3].Value = "Stock Date";
            worksheet.Cells[1, 4].Value = "Stock Before";
            worksheet.Cells[1, 5].Value = "Stock After";
            worksheet.Cells[1, 6].Value = "Month";
            worksheet.Cells[1, 7].Value = "Year";
            worksheet.Cells[1, 8].Value = "Created Date";
            worksheet.Cells[1, 9].Value = "Created By";

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].ProductName;
                worksheet.Cells[i + 2, 2].Value = data[i].StoreName;
                worksheet.Cells[i + 2, 3].Value = data[i].StockDate != null ? Convert.ToDateTime(data[i].StockDate).ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cells[i + 2, 4].Value = data[i].StockBefore;
                worksheet.Cells[i + 2, 5].Value = data[i].StockAfter;
                worksheet.Cells[i + 2, 6].Value = data[i].Month;
                worksheet.Cells[i + 2, 7].Value = data[i].Year;
                worksheet.Cells[i + 2, 8].Value = data[i].CreatedDate != null ? Convert.ToDateTime(data[i].CreatedDate).ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cells[i + 2, 9].Value = data[i].CreatedBy;
            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }

    }
}
