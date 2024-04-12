using System;
using System.Collections.Generic;
using System.Text;
using System.Security.Cryptography;
using System.Linq;
using System.Net;
using System.Data;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Collections;

public class Program
{
    static string URL = "https://xxxxx.api.netsuite.com/app/site/hosting/restlet.nl?script=xxxx&deploy=x&searchid=";
    private const string oAuthConsumerKey = "xxxx";
    private const string oAuthConsumerSecret = "xxxx";
    private const string oAuthToken = "xxxxx";
    private const string oAuthTokenSecret = "xxxxxxx";
    private const string oAuthRealm = "xxx";
    private const string oAuthVersion = "1.0";
    private const string oAuthSigMethod = "HMAC-SHA256";
    private const string netsuiteReportIds = "xxx";
    private const string destinationFolder = "C:/Users/xx/xxx/xxx";

    public static void Main()
    {
        int[] arrayNetsuiteReportIds = netsuiteReportIds.Split(',').Select(int.Parse).ToArray();
        for (int i = 0; i < arrayNetsuiteReportIds.Length; i++)
        {
            int netsuiteReportId;
            netsuiteReportId = arrayNetsuiteReportIds[i];
            string customReportId = string.Empty;
            customReportId = "customsearch" + netsuiteReportId;
            URL = URL + customReportId;
            var data = new Dictionary<string, string>()
            {
                {"script", "xxx" },
                {"deploy", "xx" },
                {"searchid", customReportId }
              //  {"startRange","0" },
               // {"endRange","1000" }
            };
            Console.WriteLine("API End Point : " + URL);
            var authString = GetAuthorizationHeader("GET", new Uri(URL), data);
            Console.WriteLine("API Authorization Header : " + authString);
           // URL = URL + "&startRange=0&endRange=1000";
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(URL);
            httpWebRequest.Method = "GET";
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Headers.Add(HttpRequestHeader.Authorization, authString);
            try
            {
                Console.WriteLine("API End Point  : " + URL);
                var response = httpWebRequest.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream());
                string responseText = reader.ReadToEnd();
                var objects = responseText.Split('}').Where(s => s.Contains(":")).Select(s => s + "}").ToList();
                Console.WriteLine("Response for :" + customReportId);

                if (objects[0].Split(':')[1].Replace("[]}", "").Length > 0)
                {
                    using (var writer = new StreamWriter(destinationFolder + netsuiteReportId + "_netsuitereport_" + GenerateTimestamp() + ".csv"))
                    {
                        List<string> headerlist = new List<string>();
                        foreach (string obj in objects)
                        {
                            string[] array;
                            List<string> rowdatalist = new List<string>();
                            List<string> csvrowdatalist = new List<string>();
                            // obj = obj.Substring(1);
                            if (obj.Contains("searchresults_row") == true)
                            {

                                string firstvalues = (string)obj.Replace("\",\"", "^").Replace("\"", "").Replace("{", "").Replace("}", "").Replace("[searchresults_row:[", "");
                                array = firstvalues.Split("^");
                                foreach (string str in array)
                                {
                                    if (str.Contains(":") == true)
                                    {
                                        Console.WriteLine(str.Substring(0, str.IndexOf(":")) + ":" + str.Substring(str.IndexOf(":") + 1));
                                        headerlist.Add(str.Substring(0, str.IndexOf(":")));
                                        rowdatalist.Add(str.Substring(str.IndexOf(":") + 1));

                                    }
                                }
                                foreach (var col in rowdatalist)
                                {
                                    if (col.Contains(",") == true)
                                    {
                                        int index = rowdatalist.FindIndex(s => s == col);
                                        string replaceValue = col.Replace(",", "!$");
                                        csvrowdatalist.Add(replaceValue);
                                    }
                                    else
                                    {
                                        csvrowdatalist.Add(col);
                                    }
                                }
                                writer.WriteLine(string.Join(",", headerlist.ToArray()));
                                writer.WriteLine(string.Join(",", csvrowdatalist.ToArray()));
                            }
                            else
                            {
                                string values = (string)obj.Replace("\",\"", "^").Replace("\"", "").Replace("{", "").Replace("}", "").Substring(1);
                                array = values.Split("^");
                                foreach (string str in array)
                                {

                                    if (str.Contains(":") == true)
                                    {
                                        Console.WriteLine(str.Substring(0, str.IndexOf(":")) + ":" + str.Substring(str.IndexOf(":") + 1));
                                        rowdatalist.Add(str.Substring(str.IndexOf(":") + 1));
                                    }
                                }
                                foreach (var col in rowdatalist.ToList())
                                {
                                    if (col.Contains(",") == true)
                                    {
                                        int index = rowdatalist.FindIndex(s => s == col);
                                        string replaceValue = col.Replace(",", "!$");
                                        csvrowdatalist.Add(replaceValue);
                                    }
                                    else
                                    {
                                        csvrowdatalist.Add(col);
                                    }
                                }
                                writer.WriteLine(string.Join(",", csvrowdatalist.ToArray()));
                            }
                        }
                    }
                    Console.WriteLine("! Report File Saved ! :" + destinationFolder);
                    Console.WriteLine("! Report File Downloaded Successfully !");
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

            URL = URL.Replace(customReportId, "");

        }
    }

    public static string GetAuthorizationHeader(string httpMethod, Uri uri, Dictionary<string, string> parameters)
    {

        if (parameters == null) parameters = new Dictionary<string, string>();
        var prms = new Dictionary<string, string>(parameters);
        prms["oauth_token"] = oAuthToken;
        prms["oauth_nonce"] = NounceGenerator(11);
        prms["oauth_timestamp"] = GenerateTimestamp();
        prms["oauth_version"] = oAuthVersion;
        prms["oauth_consumer_key"] = oAuthConsumerKey;
        prms["oauth_signature_method"] = oAuthSigMethod;
        prms["oauth_signature"] = GetSignature(httpMethod, uri, prms);
        return "OAuth realm=\""
        + Uri.EscapeDataString(oAuthRealm.ToUpperInvariant()) + "\", "
        + string.Join(", ", prms
        .Where(a => a.Key.StartsWith("oauth_"))
        .Select(kvp => Uri.EscapeDataString(kvp.Key) + "=\"" + Uri.EscapeDataString(kvp.Value) + "\"")
        );
    }
   /* public static string GetAuthorizationHeader(string httpMethod, Uri uri, Dictionary<string, string> parameters)
    {

        if (parameters == null) parameters = new Dictionary<string, string>();
        var prms = new Dictionary<string, string>(parameters);
        prms["oauth_consumer_key"] = oAuthConsumerKey;
        prms["oauth_token"] = oAuthToken;
        prms["oauth_signature_method"] = oAuthSigMethod;
        prms["oauth_timestamp"] = GenerateTimestamp();
        prms["oauth_nonce"] = NounceGenerator(11);       
        prms["oauth_version"] = oAuthVersion;
        prms["oauth_signature"] = GetSignature(httpMethod, uri, prms);
        return "OAuth realm=\""
        + Uri.EscapeDataString(oAuthRealm.ToUpperInvariant()) + "\", "
        + string.Join(", ", prms
        .Where(a => a.Key.StartsWith("oauth_"))
        .Select(kvp => Uri.EscapeDataString(kvp.Key) + "=\"" + Uri.EscapeDataString(kvp.Value) + "\"")
        );
    } */


    private static string GetSignature(string httpMethod, Uri uri, Dictionary<string, string> parameters)
    {
        string prms = string.Join("&",
        parameters.OrderBy(a => a.Key)
        .Select(kvp => string.Format("{0}={1}",
        Uri.EscapeDataString(kvp.Key),
        Uri.EscapeDataString(kvp.Value))));
        string baseString = string.Format("{0}&{1}&{2}",
        httpMethod,
        Uri.EscapeDataString(uri.AbsoluteUri.Split('?').First()), // without Query Parameters
        Uri.EscapeDataString(prms));
        using (var hmacSha256 = new HMACSHA256(Encoding.UTF8.GetBytes(string.Format("{0}&{1}",
        oAuthConsumerSecret,
        oAuthTokenSecret))))
        {
            return Convert.ToBase64String(hmacSha256.ComputeHash(Encoding.UTF8.GetBytes(baseString)));
        }
    }
    private static string GenerateTimestamp() { return ((int)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalSeconds).ToString(); }
    private static string NounceGenerator(int length)
    {
        const string chars = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return new string(Enumerable.Repeat(chars, length)
        .Select(s => s[new Random().Next(s.Length)]).ToArray());
    }
}
