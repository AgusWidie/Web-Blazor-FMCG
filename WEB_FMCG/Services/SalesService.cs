using OfficeOpenXml;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.RegularExpressions;
using WEB_FMCG.Auth;
using WEB_FMCG.Help;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface ISalesService
    {
        Task<PaginatedResult<VSale>> GetPaged(int pageNumber, int pageSize);
        Task<List<VSale>> GetAll();
        Task<List<VSale>> GetAllSalesTL();
        Task<List<VSale>> GetAllSalesNonTL();
        Task<string> CheckUserName(string UserName);
        Task Create(Sale a);
        Task Update(Sale a);
        byte[] GenerateExportSales(List<VSale> data);
    }

    public class SalesService : ISalesService
    {
        private readonly HttpClient _http;
        private readonly AuthService _auth;
        private string baseAPI = "master/api/v1/sales";
        public SalesService(HttpClient http, AuthService auth)
        {
            _http = http;
            _auth = auth;
        }

        public async Task<PaginatedResult<VSale>> GetPaged(int pageNumber, int pageSize)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    PropertyNameCaseInsensitive = true
                };
                options.Converters.Add(new DateOnlyJsonConverter());
                var token = await _auth.GetTokenAsync();
                //Console.WriteLine("Token : " + token);

                if (string.IsNullOrWhiteSpace(token))
                {
                    throw new InvalidOperationException("User not logged in");
                }

                var url = $"{baseAPI}/GetPagedSales?pageNumber={pageNumber}&pageSize={pageSize}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                // Baca response JSON sebagai string
                //var jsonString = await response.Content.ReadAsStringAsync();

                // Print JSON
                //Console.WriteLine("Response : " + jsonString);

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<PaginatedResult<VSale>>>(options);
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


        public async Task<List<VSale>> GetAll()
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

                var request = new HttpRequestMessage(HttpMethod.Get, $"{baseAPI}/GetAllSales");
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<List<VSale>>>(options);
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

        public async Task<List<VSale>> GetAllSalesTL()
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

                var request = new HttpRequestMessage(HttpMethod.Get, $"{baseAPI}/GetSalesIsTL");
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<List<VSale>>>(options);
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

        public async Task<List<VSale>> GetAllSalesNonTL()
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

                var request = new HttpRequestMessage(HttpMethod.Get, $"{baseAPI}/GetSalesNonIsTL");
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await _http.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var apiResponse = await response.Content.ReadFromJsonAsync<ApiResponse<List<VSale>>>(options);
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

        public async Task<string> CheckUserName(string UserName)
        {
            // Definisi pola: huruf kecil, angka, dan titik
            string pattern = @"^[a-z0-9.]+$";

            if (string.IsNullOrEmpty(UserName))
            {
                //throw new Exception("Username Cannot Be Empty.");
                return "Failed : Username Cannot Be Empty.";
            }

            if (!Regex.IsMatch(UserName, pattern))
            {
                //throw new Exception("Format username tidak valid. Use only lowercase letters, numbers, and periods.");
                return "Failed : Format username tidak valid. Use only lowercase letters, numbers, and periods.";
            }

            return "Success : Check User Name.";
        }

        public async Task Create(Sale a)
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

                var request = new HttpRequestMessage(HttpMethod.Post, $"{baseAPI}/CreateSales")
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

        public async Task Update(Sale a)
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

                var request = new HttpRequestMessage(HttpMethod.Put, $"{baseAPI}/UpdateSales")
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

        public byte[] GenerateExportSales(List<VSale> data)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Sales");

            // Header
            worksheet.Cells[1, 1].Value = "Sales ID";
            worksheet.Cells[1, 2].Value = "Sales Name";
            worksheet.Cells[1, 3].Value = "Role ID";
            worksheet.Cells[1, 4].Value = "Role Name";
            worksheet.Cells[1, 5].Value = "Gender";
            worksheet.Cells[1, 6].Value = "Religion";
            worksheet.Cells[1, 7].Value = "Education";
            worksheet.Cells[1, 8].Value = "Address";
            worksheet.Cells[1, 9].Value = "Place Of Birth";
            worksheet.Cells[1, 10].Value = "Date Of Birth";
            worksheet.Cells[1, 11].Value = "Email";
            worksheet.Cells[1, 12].Value = "Telephone";
            worksheet.Cells[1, 13].Value = "User Name";

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].SalesId;
                worksheet.Cells[i + 2, 2].Value = data[i].SalesName;
                worksheet.Cells[i + 2, 3].Value = data[i].RoleId;
                worksheet.Cells[i + 2, 4].Value = data[i].RoleName;
                worksheet.Cells[i + 2, 5].Value = data[i].Gender;
                worksheet.Cells[i + 2, 6].Value = data[i].Religion;
                worksheet.Cells[i + 2, 7].Value = data[i].Education;
                worksheet.Cells[i + 2, 8].Value = data[i].Address;
                worksheet.Cells[i + 2, 9].Value = data[i].PlaceOfBirth;
                worksheet.Cells[i + 2, 10].Value = data[i].DateOfBirth != null ? Convert.ToDateTime(data[i].DateOfBirth).ToString("yyyy-MM-dd") : "";
                worksheet.Cells[i + 2, 11].Value = data[i].Email;
                worksheet.Cells[i + 2, 12].Value = data[i].Telephone;
                worksheet.Cells[i + 2, 13].Value = data[i].UserName;
            }

            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }
}
