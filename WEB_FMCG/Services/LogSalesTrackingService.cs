using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface ILogSalesTrackingService
    {
        Task<PaginatedResult<VSalesTracking>> GetPaged(int pageNumber, int pageSize, int month, int year);
        Task<IEnumerable<VSalesTracking>> SalesTrackingExport(int month, int year);
        byte[] GenerateExportSalesTracking(List<VSalesTracking> data);
    }

    public class LogSalesTrackingService : ILogSalesTrackingService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "realtime/api/v1/SalesTracking";
        public LogSalesTrackingService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<VSalesTracking>> GetPaged(int pageNumber, int pageSize, int month, int year)
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

                var url = $"{baseAPI}/GetPagedSalesTracking?pageNumber={pageNumber}&pageSize={pageSize}&month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<VSalesTracking>>>(options);
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

        public async Task<IEnumerable<VSalesTracking>> SalesTrackingExport(int month, int year)
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

                var url = $"{baseAPI}/GetSalesTrackingImport?month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<IEnumerable<VSalesTracking>>>(options);
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

        public byte[] GenerateExportSalesTracking(List<VSalesTracking> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("SalesTracking");

            // Header
            worksheet.Cells[1, 1].Value = "Sales Id";
            worksheet.Cells[1, 2].Value = "Sales Name";
            worksheet.Cells[1, 3].Value = "Gender";
            worksheet.Cells[1, 4].Value = "Status";
            worksheet.Cells[1, 5].Value = "Religion";
            worksheet.Cells[1, 6].Value = "Education";
            worksheet.Cells[1, 7].Value = "Address";
            worksheet.Cells[1, 8].Value = "Email";
            worksheet.Cells[1, 9].Value = "Telephone";
            worksheet.Cells[1, 10].Value = "Activity";
            worksheet.Cells[1, 11].Value = "Longitude";
            worksheet.Cells[1, 12].Value = "Latitude";
            worksheet.Cells[1, 13].Value = "Month";
            worksheet.Cells[1, 14].Value = "Year";

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].SalesId;
                worksheet.Cells[i + 2, 2].Value = data[i].SalesName;
                worksheet.Cells[i + 2, 3].Value = data[i].Gender;
                worksheet.Cells[i + 2, 4].Value = data[i].Status;
                worksheet.Cells[i + 2, 5].Value = data[i].Religion;
                worksheet.Cells[i + 2, 6].Value = data[i].Education;
                worksheet.Cells[i + 2, 7].Value = data[i].Address;
                worksheet.Cells[i + 2, 8].Value = data[i].Email;
                worksheet.Cells[i + 2, 9].Value = data[i].Telephone;
                worksheet.Cells[i + 2, 10].Value = data[i].Activity;
                worksheet.Cells[i + 2, 11].Value = data[i].Longitude;
                worksheet.Cells[i + 2, 12].Value = data[i].Latitude;
                worksheet.Cells[i + 2, 13].Value = data[i].Month;
                worksheet.Cells[i + 2, 14].Value = data[i].Year;
                worksheet.Cells[i + 2, 15].Value = data[i].CreatedDate != null ? Convert.ToDateTime(data[i].CreatedDate).ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cells[i + 2, 16].Value = data[i].CreatedBy;
            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }
}
