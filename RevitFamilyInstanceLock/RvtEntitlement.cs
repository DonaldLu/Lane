using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Specialized;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;

namespace RevitFamilyInstanceLock
{
    public class RvtEntitlement
    {
        [Serializable]
        public class EntitlementResponse
        {
            public string UserId
            {
                get;
                set;
            }

            public string AppId
            {
                get;
                set;
            }

            public bool IsValid
            {
                get;
                set;
            }

            public string Message
            {
                get;
                set;
            }
        }

        public const string _baseApiUrl = "https://apps.autodesk.com/";

        public bool IsEntitledAsync(string appId, string userId)
        {
            new RvtEntitlement.EntitlementResponse();
            bool result2;
            using (HttpClient httpClient = new HttpClient())
            {
                UriBuilder expr_16 = new UriBuilder("https://apps.autodesk.com/webservices/checkentitlement");
                expr_16.Port = -1;
                NameValueCollection nameValueCollection = HttpUtility.ParseQueryString(expr_16.Query);
                nameValueCollection["userid"] = userId;
                nameValueCollection["appid"] = appId;
                expr_16.Query = nameValueCollection.ToString();
                string text = expr_16.ToString();
                httpClient.DefaultRequestHeaders.Accept.Clear();
                httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                string result = httpClient.GetStringAsync(text).Result;
                if (RvtEntitlement.IsValidJson(result))
                {
                    result2 = JsonConvert.DeserializeObject<RvtEntitlement.EntitlementResponse>(result).IsValid;
                }
                else
                {
                    result2 = false;
                }
            }
            return result2;
        }

        private static bool IsValidJson(string strInput)
        {
            if (string.IsNullOrWhiteSpace(strInput))
            {
                return false;
            }
            strInput = strInput.Trim();
            if ((strInput.StartsWith("{") && strInput.EndsWith("}")) || (strInput.StartsWith("[") && strInput.EndsWith("]")))
            {
                try
                {
                    JToken.Parse(strInput);
                    bool result = true;
                    return result;
                }
                catch (JsonReaderException arg_52_0)
                {
                    Console.WriteLine(arg_52_0.Message);
                    bool result = false;
                    return result;
                }
                catch (Exception arg_60_0)
                {
                    Console.WriteLine(arg_60_0.ToString());
                    bool result = false;
                    return result;
                }
                return false;
            }
            return false;
        }
    }
}