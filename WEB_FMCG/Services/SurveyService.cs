using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using OfficeOpenXml;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{

    public interface ISurveyService
    {
        Task<PaginatedResult<Survey>> GetPaged(int pageNumber, int pageSize);
        Task<List<Survey>> GetAll();
        Task Create(Survey a);
        Task Update(Survey a);
        byte[] GenerateExportSurvey(List<Survey> data);
    }

    public class SurveyService : ISurveyService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "master/api/v1/Survey";
        public SurveyService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<Survey>> GetPaged(int pageNumber, int pageSize)
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

                var url = $"{baseAPI}/GetPagedSurvey?pageNumber={pageNumber}&pageSize={pageSize}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<Survey>>>(options);
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

        public async Task<List<Survey>> GetAll()
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

                var request = new HttpRequestMessage(HttpMethod.Get, $"{baseAPI}/GetAllSurvey");
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);


                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<List<Survey>>>(options);
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

        public async Task Create(Survey a)
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

                var request = new HttpRequestMessage(HttpMethod.Post, $"{baseAPI}/CreateSurvey")
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

        public async Task Update(Survey a)
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

                var request = new HttpRequestMessage(HttpMethod.Put, $"{baseAPI}/UpdateSurvey")
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

        public byte[] GenerateExportSurvey(List<Survey> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Survey");

            // Header
            worksheet.Cells[1, 1].Value = "Survey ID";
            worksheet.Cells[1, 2].Value = "Title";
            worksheet.Cells[1, 3].Value = "Start Date";
            worksheet.Cells[1, 4].Value = "End Date";
            worksheet.Cells[1, 5].Value = "Description";
            worksheet.Cells[1, 6].Value = "Status";

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].SurveyId;
                worksheet.Cells[i + 2, 2].Value = data[i].Title;
                worksheet.Cells[i + 2, 3].Value = data[i].StartDate != null ? Convert.ToDateTime(data[i].StartDate).ToString("yyyy-MM-dd") : "";
                worksheet.Cells[i + 2, 4].Value = data[i].EndDate != null ? Convert.ToDateTime(data[i].EndDate).ToString("yyyy-MM-dd") : "";
                worksheet.Cells[i + 2, 5].Value = data[i].Description;
                worksheet.Cells[i + 2, 6].Value = data[i].Status;
            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }

}
