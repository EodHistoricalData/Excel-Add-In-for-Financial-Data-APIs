using EODAddIn.Model;

using Newtonsoft.Json;

namespace EODAddIn.Utils
{
    public static class APIEOD
    {
        /// <summary>
        /// Запрос на получение информации о пользователе
        /// </summary>
        /// <param name="api_token"></param>
        /// <returns></returns>
        public static User User(string api_token)
        {
            string url = "https://eodhistoricaldata.com/api/user";
            string s = Response.GET(url, "api_token=" + api_token);
            return JsonConvert.DeserializeObject<User>(s);
        }
    }
}
