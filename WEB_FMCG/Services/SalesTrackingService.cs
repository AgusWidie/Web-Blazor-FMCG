using Microsoft.AspNetCore.SignalR.Client;
using WEB_FMCG.Models;

namespace WEB_FMCG.Services
{
    public interface ISalesTrackingService
    {
        Task StartAsync();
    }

    public class SalesTrackingService : ISalesTrackingService
    {
        private HubConnection _hubConnection;

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
        public event Action<VSalesTracking> OnSalesReceived;
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.

        public async Task StartAsync()
        {
            var baseUrl = "https://api.coresight-hub.cloud/sales-tracking";
            var connection = new HubConnectionBuilder()
                .WithUrl($"{baseUrl}/salesTrackingHub")
                .WithAutomaticReconnect()
                .Build();

            _hubConnection.On<VSalesTracking>("ReceiveSalesTracking", (data) =>
            {
                OnSalesReceived?.Invoke(data);
            });

            await _hubConnection.StartAsync();
        }
    }
}
