using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface ISurveyQuestionService
    {
        Task<PaginatedResult<VSurveyQuestion>> GetPaged(int pageNumber, int pageSize, int month, int year);
        Task<IEnumerable<VSurveyQuestion>> SurveyQuestionExport(int month, int year);
        byte[] GenerateExportSurveyQuestion(List<VSurveyQuestion> data);
    }

    public class SurveyQuestionService : ISurveyQuestionService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "transaction/api/v1/SurveyQuestion";
        public SurveyQuestionService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<VSurveyQuestion>> GetPaged(int pageNumber, int pageSize, int month, int year)
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

                var url = $"{baseAPI}/GetPagedSurveyQuestion?pageNumber={pageNumber}&pageSize={pageSize}&month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<VSurveyQuestion>>>(options);
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

        public async Task<IEnumerable<VSurveyQuestion>> SurveyQuestionExport(int month, int year)
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

                var url = $"{baseAPI}/GetSurveyQuestionExport?month={month}&year={year}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<IEnumerable<VSurveyQuestion>>>(options);
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

        public byte[] GenerateExportSurveyQuestion(List<VSurveyQuestion> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("SurveyQuestion");

            // Header
            worksheet.Cells[1, 1].Value = "Survey Question Id";
            worksheet.Cells[1, 2].Value = "Survey Id";
            worksheet.Cells[1, 3].Value = "Title";
            worksheet.Cells[1, 4].Value = "Description";
            worksheet.Cells[1, 5].Value = "Start Date";
            worksheet.Cells[1, 6].Value = "End Date";
            worksheet.Cells[1, 5].Value = "Question Text";
            worksheet.Cells[1, 6].Value = "Question Type";
            worksheet.Cells[1, 7].Value = "Options";
            worksheet.Cells[1, 8].Value = "Created Date";
            worksheet.Cells[1, 9].Value = "Created By";

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].SurveyQuestionId;
                worksheet.Cells[i + 2, 2].Value = data[i].SurveyId;
                worksheet.Cells[i + 2, 3].Value = data[i].Title;
                worksheet.Cells[i + 2, 4].Value = data[i].Description;
                worksheet.Cells[i + 2, 5].Value = data[i].StartDate != null ? Convert.ToDateTime(data[i].StartDate).ToString("yyyy-MM-dd") : "";
                worksheet.Cells[i + 2, 6].Value = data[i].EndDate != null ? Convert.ToDateTime(data[i].EndDate).ToString("yyyy-MM-dd") : "";
                worksheet.Cells[i + 2, 7].Value = data[i].QuestionText;
                worksheet.Cells[i + 2, 8].Value = data[i].QuestionType;
                worksheet.Cells[i + 2, 9].Value = data[i].Options;
                worksheet.Cells[i + 2, 10].Value = data[i].CreatedDate != null ? Convert.ToDateTime(data[i].CreatedDate).ToString("yyyy-MM-dd") : "";
                worksheet.Cells[i + 2, 11].Value = data[i].CreatedBy;
            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }
}
