using System;

namespace EODAddIn.Model
{
    public class User
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string SubscriptionType { get; set; }
        public string PaymentMethod { get; set; }
        public int ApiRequests { get; set; }
        public DateTime ApiRequestsDate { get; set; }
        public int DailyRateLimit { get; set; }
        public string InviteToken { get; set; }
        public int InviteTokenClicked { get; set; }

        public User()
        {

        }

    }
}
