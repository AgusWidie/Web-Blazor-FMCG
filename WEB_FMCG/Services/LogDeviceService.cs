using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Runtime.CompilerServices;
using System.Text.Json;
using WEB_FMCG.Auth;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface ILogDeviceService
    {
        Task<PaginatedResult<LoginDevice>> GetPaged(int pageNumber, int pageSize, int month, int year);
        Task<List<LoginDevice>> GetLoginDevice(int month, int year);
        byte[] GenerateExportLogDevice(List<LoginDevice> data);
    }

    public class LogDeviceService : ILogDeviceService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "master/api/v1/LogDevice";
        public LogDeviceService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<LoginDevice>> GetPaged(int pageNumber, int pageSize, int month, int year)
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

                var url = $"{baseAPI}/GetPagedLoginDevice?pageNumber={pageNumber}&pageSize={pageSize}&month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<LoginDevice>>>(options);
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

        public async Task<List<LoginDevice>> GetLoginDevice(int month, int year)
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

                var url = $"{baseAPI}/GetLoginDevice?month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<List<LoginDevice>>>(options);
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

        public byte[] GenerateExportLogDevice(List<LoginDevice> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("LogDevice");

            // Header
            worksheet.Cells[1, 1].Value = "Sales ID";
            worksheet.Cells[1, 2].Value = "Device ID";
            worksheet.Cells[1, 3].Value = "Login Date";
            worksheet.Cells[1, 4].Value = "Is Login";
            worksheet.Cells[1, 5].Value = "Logout Date";

            string LoginDate = "-";
            string LogoutDate = "-";

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].SalesId;
                worksheet.Cells[i + 2, 2].Value = data[i].DeviceId;
                if(data[i].LoginDate != null)
                {
                    LoginDate = Convert.ToDateTime(data[i].LoginDate).ToString("yyyy-MM-dd HH:mm:ss");
                }
                if (data[i].LogoutDate != null)
                {
                    LogoutDate = Convert.ToDateTime(data[i].LogoutDate).ToString("yyyy-MM-dd HH:mm:ss");
                }
                worksheet.Cells[i + 2, 3].Value = LoginDate;
                worksheet.Cells[i + 2, 4].Value = data[i].IsLogin;
                worksheet.Cells[i + 2, 5].Value = LogoutDate;
               
            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }
}
