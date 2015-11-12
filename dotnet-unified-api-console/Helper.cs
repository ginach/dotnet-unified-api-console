#region

using System;
using System.IO;
using System.Threading.Tasks;
using System.Net.Http;
#endregion

namespace MicrosoftGraphSampleConsole
{
    internal class Helper
    {
        /// <summary>
        ///     Returns a random string of upto 32 characters.
        /// </summary>
        /// <returns>String of upto 32 characters.</returns>
        public static string GetRandomString(int length = 32)
        {
            //because GUID can't be longer than 32
            return Guid.NewGuid().ToString("N").Substring(0, length > 32 ? 32 : length);
        }

        public static async Task<Stream> GetRestRequestStream(string restRequest, string token)
        {

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, Constants.ResourceUrl.TrimEnd('/') + restRequest))
                {
                    request.Headers.Add("Authorization", "Bearer " + token);
                    using (var response = await client.SendAsync(request))
                    {
                        var content = await response.Content.ReadAsStreamAsync();
                        if (response.IsSuccessStatusCode)
                        {
                            return content;
                        }
                        else
                        {
                            throw new Exception();
                        }
                    }
                }
            }

        }
    }
}