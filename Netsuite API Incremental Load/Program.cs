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
    //static string URL = "https://5233349-sb2.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=1172&deploy=1&searchid=";
   // static string URL = "https://5233349.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=1290&deploy=1&searchid=";
    static string URL = "https://5233349-sb2.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=1290&deploy=1&searchid=";
    //private const string oAuthConsumerKey = "50fd5283e0037b590dff49a4eaab211c7b1bcd8c8f5b5fabad59626238672a1a";
    //private const string oAuthConsumerSecret = "a98a902e3d4a6ad10fbea763593f04c37e0528a44274bef80c3570766c97bd0c";
    //private const string oAuthToken = "5ba0963065f6aaaf1c1d4de6ddb6b7eda0e3db61e7fa854864c21f0fa97ea9cf";
    //private const string oAuthTokenSecret = "94816d9cd37a788bb8cb37a80252a6a5f1d37f1ade24aa4308419271005f81a5";
    private const string oAuthConsumerKey = "b4f18bf098a5f18d74a1840dfa0bcf0cbaa78d0e3f4eb07c174eb147864c5cb2";
    private const string oAuthConsumerSecret = "ba46bddd14f12ce6dc985041309f1ad23af8517767aef3cc0d52bac640ea4dda";
    private const string oAuthToken = "4b9948049dc11442fd4b0fe0e40afe3076a594f6a982c737350f426b8d6f3530";
    private const string oAuthTokenSecret = "151b9abadb289b0d6f4feee1dd9b8f096a0a20c0743eb8c4ec49afe6051d7168";
    //private const string oAuthRealm = "5233349";
    private const string oAuthRealm = "5233349_SB2";
    private const string oAuthVersion = "1.0";
    private const string oAuthSigMethod = "HMAC-SHA256";
    private const string netsuiteReportIds = "2437";
    private const string destinationFolder = "C:/Users/40000450/OneDrive - Buck Global LLC/Desktop/Sheetal/Development/Netsuite/Downloads/";

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
                {"script", "1290" },
                {"deploy", "1" },
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