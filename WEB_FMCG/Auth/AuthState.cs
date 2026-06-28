namespace WEB_FMCG.Auth
{
    public class AuthState
    {
        // token disimpan di memori
        public string? AccessToken { get; private set; }
        public string? UserName { get; private set; }
        public string? RoleId { get; private set; }

        public bool IsAuthenticated => !string.IsNullOrEmpty(AccessToken);

        // Bisa simpan semua sekaligus
        public void SetToken(string token, string userName, string roleId)
        {
            AccessToken = token;
            UserName = userName;
            RoleId = roleId;
        }

        public void Clear()
        {
            AccessToken = null;
            UserName = null;
            RoleId = null;
        }
    }
}
