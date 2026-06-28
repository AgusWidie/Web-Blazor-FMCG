namespace WEB_FMCG.DTO
{
  
    public class ApiLoginResponse
    {
        public string? Message { get; set; }
        public int? StatusCode { get; set; }
        public bool? Success { get; set; }
        public LoginResponse? Data { get; set; }
    }

    public class LoginResponse
    {
        public string? SalesId { get; set; }
        public string? SalesName { get; set; }
        public string? UserName { get; set; }
        public string? RoleId { get; set; }
        public string? Token { get; set; }
    }
}
