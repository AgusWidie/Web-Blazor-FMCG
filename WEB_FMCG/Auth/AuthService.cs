using Blazored.LocalStorage;

public class AuthSession
{
    public string SalesId { get; set; } = default!;
    public string UserName { get; set; } = default!;
    public string RoleId { get; set; } = default!;
    public string AccessToken { get; set; } = default!;
}

public class AuthService
{
    private readonly ILocalStorageService _storage;

    private const string SessionsKey = "auth_sessions";
    private const string ActiveUserKey = "active_user";

    public Dictionary<string, AuthSession> Sessions { get; private set; }
        = new();

    public string? ActiveUserId { get; private set; }

    public AuthSession? CurrentSession =>
        ActiveUserId != null && Sessions.ContainsKey(ActiveUserId)
            ? Sessions[ActiveUserId]
            : null;

    public bool IsAuthenticated => CurrentSession != null;

    public AuthService(ILocalStorageService storage)
    {
        _storage = storage;
    }

    // WAJIB dipanggil saat app start
    public async Task InitializeAsync()
    {
        Sessions = await _storage.GetItemAsync<Dictionary<string, AuthSession>>(SessionsKey)
                   ?? new();

        ActiveUserId = await _storage.GetItemAsync<string>(ActiveUserKey);

        //Console.WriteLine("Sessions count: " + Sessions.Count);
        //Console.WriteLine("ActiveUserId: " + ActiveUserId);

        if (ActiveUserId != null && Sessions.ContainsKey(ActiveUserId))
        {
            //Console.WriteLine("TOKEN DI STORAGE: " + Sessions[ActiveUserId].AccessToken);
        }
    }

    public async Task<string?> GetTokenAsync()
    {
        // Ambil active user id dulu (source of truth)
        var activeUserId = ActiveUserId
            ?? await _storage.GetItemAsync<string>(ActiveUserKey);

        if (string.IsNullOrEmpty(activeUserId))
            return null;

        // 2️⃣ Fallback ke storage
        var sessions = Sessions
            ?? await _storage.GetItemAsync<Dictionary<string, AuthSession>>(SessionsKey);

        if (sessions == null || !sessions.TryGetValue(activeUserId, out var session))
            return null;

        // 3️⃣ Sync balik ke memory
        Sessions = sessions;
        ActiveUserId = activeUserId;

        return session.AccessToken;
    }

    public async Task<string?> GetUserNameAsync()
    {
        // Ambil active user id dulu (source of truth)
        var activeUserId = ActiveUserId
            ?? await _storage.GetItemAsync<string>(ActiveUserKey);

        if (string.IsNullOrEmpty(activeUserId))
            return null;

        // 2️⃣ Fallback ke storage
        var sessions = Sessions
            ?? await _storage.GetItemAsync<Dictionary<string, AuthSession>>(SessionsKey);

        if (sessions == null || !sessions.TryGetValue(activeUserId, out var session))
            return null;

        // 3️⃣ Sync balik ke memory
        Sessions = sessions;
        ActiveUserId = activeUserId;

        return session.UserName;
    }

    // Login / tambah user
    public async Task SetTokenAsync(
        string salesId,
        string userName,
        string token,
        string roleId)
    {
        Sessions[salesId] = new AuthSession
        {
            SalesId = salesId,
            UserName = userName,
            RoleId = roleId,
            AccessToken = token
        };

        ActiveUserId = salesId;

        await PersistAsync();
    }

    // Pindah user aktif
    public async Task SwitchUserAsync(string salesId)
    {
        if (!Sessions.ContainsKey(salesId))
            throw new InvalidOperationException("User tidak ditemukan");

        ActiveUserId = salesId;
        await _storage.SetItemAsync(ActiveUserKey, salesId);
    }

    // Logout user tertentu
    public async Task LogoutAsync(string userId)
    {
        if (Sessions.Remove(userId))
        {
            if (ActiveUserId == userId)
                ActiveUserId = null;

            await PersistAsync();
        }
    }

    
    // Logout semua user
    public async Task LogoutAllAsync()
    {
        Sessions.Clear();
        ActiveUserId = null;

        await _storage.RemoveItemAsync(SessionsKey);
        await _storage.RemoveItemAsync(ActiveUserKey);
    }

    private async Task PersistAsync()
    {
        await _storage.SetItemAsync(SessionsKey, Sessions);
        await _storage.SetItemAsync(ActiveUserKey, ActiveUserId);
    }
}
