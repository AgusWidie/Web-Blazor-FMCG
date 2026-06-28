using Microsoft.AspNetCore.Components.WebAssembly;

namespace WEB_FMCG.Help
{
    public class AppState
    {
        public string? _salesId { get; private set; }
        public string? _roleId { get; private set; }


        public void SetUser(string salesId, string roleId)
        {
            _salesId = salesId;
            _roleId = roleId;
        }

        public void Clear()
        {
            _salesId = null;
            _roleId = null;
        }
    }
}
