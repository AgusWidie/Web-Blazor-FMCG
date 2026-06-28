using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.Auth;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface ISubmissionLeaveService
    {
        Task<PaginatedResult<VSubmissionLeave>> GetPaged(int pageNumber, int month, int year);
        Task<IEnumerable<VSubmissionLeave>> SubmissionLeaveExport(int month, int year);
        Task Create(SubmissionLeave a);
        Task Update(SubmissionLeave a);
        byte[] GenerateExportSubmissionLeave(List<VSubmissionLeave> data);
    }

    public class SubmissionLeaveService : ISubmissionLeaveService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "attendance/api/v1/SubmissionLeave";
        public SubmissionLeaveService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<VSubmissionLeave>> GetPaged(int pageNumber, int month, int year)
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

                var url = $"{baseAPI}/GetPagedSubmissionLeave?pageNumber={pageNumber}&month={month}&year={year}";
                
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<VSubmissionLeave>>>(options);
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

        public async Task<IEnumerable<VSubmissionLeave>> SubmissionLeaveExport(int month, int year)
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

                var url = $"{baseAPI}/GetSubmissionLeaveExport?month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<IEnumerable<VSubmissionLeave>>>(options);
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

        public async Task Create(SubmissionLeave a)
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

                var request = new HttpRequestMessage(HttpMethod.Post, $"{baseAPI}/CreateSubmissionLeave")
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

        public async Task Update(SubmissionLeave a)
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

                var request = new HttpRequestMessage(HttpMethod.Put, $"{baseAPI}/UpdateAttendanceLeave")
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

        public byte[] GenerateExportSubmissionLeave(List<VSubmissionLeave> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("SubmissionLeave");

            // Header
            worksheet.Cells[1, 1].Value = "Leave ID";
            worksheet.Cells[1, 2].Value = "Sales ID";
            worksheet.Cells[1, 3].Value = "Sales Name";
            worksheet.Cells[1, 4].Value = "Leave Date From";
            worksheet.Cells[1, 5].Value = "Leave Date To";
            worksheet.Cells[1, 6].Value = "Total Leave";
            worksheet.Cells[1, 7].Value = "File Url Photo";
            worksheet.Cells[1, 8].Value = "Approve Date";
            worksheet.Cells[1, 9].Value = "Approve By";
            worksheet.Cells[1, 10].Value = "Reject Date";
            worksheet.Cells[1, 11].Value = "Reject By";
            worksheet.Cells[1, 12].Value = "Reason Reject";
            worksheet.Cells[1, 13].Value = "Cancel Date";
            worksheet.Cells[1, 14].Value = "Cancel By";

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].LeaveId;
                worksheet.Cells[i + 2, 2].Value = data[i].SalesId;
                worksheet.Cells[i + 2, 3].Value = data[i].SalesName;
                worksheet.Cells[i + 2, 4].Value = data[i].LeaveDateFrom;
                worksheet.Cells[i + 2, 5].Value = data[i].LeaveDateTo;
                worksheet.Cells[i + 2, 6].Value = data[i].TotalLeave;
                worksheet.Cells[i + 2, 7].Value = data[i].FileUrlPhoto;
                worksheet.Cells[i + 2, 8].Value = data[i].ApproveDate != null ? Convert.ToDateTime(data[i].ApproveDate).ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cells[i + 2, 9].Value = data[i].ApproveBy;
                worksheet.Cells[i + 2, 10].Value = data[i].RejectDate != null ? Convert.ToDateTime(data[i].RejectDate).ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cells[i + 2, 11].Value = data[i].RejectBy;
                worksheet.Cells[i + 2, 12].Value = data[i].ReasonReject;
                worksheet.Cells[i + 2, 13].Value = data[i].CancelDate != null ? Convert.ToDateTime(data[i].CancelDate).ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cells[i + 2, 13].Value = data[i].CancelBy;
            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }
}
