using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.Auth;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface IAttendanceService
    {
        Task<PaginatedResult<VAttendance>> GetPaged(int pageNumber, int month, int year);
        Task<IEnumerable<VAttendance>> AttendanceExport(int month, int year);
        Task Create(Attendance a);
        Task Update(Attendance a);
        byte[] GenerateExportAttendance(List<VAttendance> data);
    }

    public class AttendanceService : IAttendanceService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "attendance/api/v1/Attendance";
        public AttendanceService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<VAttendance>> GetPaged(int pageNumber, int month, int year)
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

                var url = $"{baseAPI}/GetPagedAttendance?pageNumber={pageNumber}&month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                //var jsonString = await response.Content.ReadAsStringAsync();

                ////Print JSON
                //Console.WriteLine("Response : " + jsonString);

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<VAttendance>>>(options);
                if (apiResponse == null || apiResponse.Data == null)
                    throw new InvalidOperationException("Data Null");

                //var json = JsonSerializer.Serialize(apiResponse.Data, new JsonSerializerOptions
                //{
                //    WriteIndented = true
                //});

                //Console.WriteLine(json);

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

        public async Task<IEnumerable<VAttendance>> AttendanceExport(int month, int year)
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

                var url = $"{baseAPI}/GetAttendanceExport?month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<IEnumerable<VAttendance>>>(options);
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

        public async Task Create(Attendance a)
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

                var request = new HttpRequestMessage(HttpMethod.Post, $"{baseAPI}/CreateAttendance")
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

        public async Task Update(Attendance a)
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

                var request = new HttpRequestMessage(HttpMethod.Put, $"{baseAPI}/UpdateAttendance")
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

        public byte[] GenerateExportAttendance(List<VAttendance> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Attendance");

            // Header
            worksheet.Cells[1, 1].Value = "Sales ID";
            worksheet.Cells[1, 2].Value = "Sales Name";
            worksheet.Cells[1, 3].Value = "Attendance In";
            worksheet.Cells[1, 4].Value = "Longitude In";
            worksheet.Cells[1, 5].Value = "Latitude In";
            worksheet.Cells[1, 6].Value = "Url Photo In";
            worksheet.Cells[1, 7].Value = "Attendance Out";
            worksheet.Cells[1, 8].Value = "Longitude Out";
            worksheet.Cells[1, 9].Value = "Latitude Out";
            worksheet.Cells[1, 10].Value = "Url Photo Out";
            worksheet.Cells[1, 11].Value = "Month";
            worksheet.Cells[1, 12].Value = "Year";
            worksheet.Cells[1, 13].Value = "Description";

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].SalesId;
                worksheet.Cells[i + 2, 2].Value = data[i].SalesName;
                worksheet.Cells[i + 2, 3].Value = data[i].IncomingAttendance != null ? Convert.ToDateTime(data[i].IncomingAttendance).ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cells[i + 2, 4].Value = data[i].IncomingLongitude;
                worksheet.Cells[i + 2, 5].Value = data[i].IncomingLatitude;
                worksheet.Cells[i + 2, 6].Value = data[i].IncomingUrlPhoto;
                worksheet.Cells[i + 2, 7].Value = data[i].RepeatAttendance != null ? Convert.ToDateTime(data[i].RepeatAttendance).ToString("yyyy-MM-dd HH:mm:ss") : "";
                worksheet.Cells[i + 2, 8].Value = data[i].RepeatLongitude;
                worksheet.Cells[i + 2, 9].Value = data[i].RepeatLatitude;
                worksheet.Cells[i + 2, 10].Value = data[i].RepeatUrlPhoto;
                worksheet.Cells[i + 2, 11].Value = data[i].Month;
                worksheet.Cells[i + 2, 12].Value = data[i].Year;
                worksheet.Cells[i + 2, 13].Value = data[i].Description;
            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }
}
