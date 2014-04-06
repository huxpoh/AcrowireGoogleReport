using System;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace Google.Apis.Sample.MVC.Utils
{
    public class AcrowireService
    {
        public AcrowireService()
        {
            Cookies = new NameValueCollection();
        }

        public string CompareCookieToHeader()
        {
            var cookiesResult = new StringBuilder();
            var a = 0;
            foreach (var key in Cookies)
            {
                if (a > 0)
                {
                    cookiesResult.Append("; ");
                }
                cookiesResult.Append(key + "=" + Cookies[key.ToString()]);
                a += 1;
            }
            return cookiesResult.ToString();
        }

        public NameValueCollection Cookies { set; get; }

        public void SetSessionCookie()
        {
            var request = (HttpWebRequest)WebRequest.Create("http://pmp.acrowire.com/public/index.php?path_info=login");
            request.Method = "GET";
            request.Accept = "Accept=text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:27.0) Gecko/20100101 Firefox/27.0";
            var response = (HttpWebResponse)request.GetResponse();
            Console.WriteLine(response.Headers["Set-Cookie"]);
            var key = new string(response.Headers["Set-Cookie"].TakeWhile(x => x != '=').ToArray());
            var value = new string(response.Headers["Set-Cookie"].Skip(key.Length + 1).TakeWhile(x => x != ';').ToArray());
            Cookies.Add(key, value);
        }


        public void SetAuthCookie(string login, string password)
        {
            var request = (HttpWebRequest)WebRequest.Create("http://pmp.acrowire.com/public/index.php?path_info=login");

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            var buffer = Encoding.ASCII.GetBytes(@"login%5Bemail%5D=" + login + "&login%5Bpassword%5D=" + password + "&submitted=submitted");
            request.ContentLength = buffer.Length;
            var requestStream = request.GetRequestStream();
            requestStream.Write(buffer, 0, buffer.Length);
            requestStream.Close();

            request.CookieContainer = new CookieContainer();

            var response = (HttpWebResponse)request.GetResponse();
            response.Close();

            foreach (Cookie cook in response.Cookies)
            {
                Cookies.Add(cook.Name, cook.Value);
            }
        }

        public string GetHtml(string cookie)
        {
            var request = (HttpWebRequest)WebRequest.Create("http://pmp.acrowire.com/public/index.php?path_info=time%2F8");
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Accept = "Accept=text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:27.0) Gecko/20100101 Firefox/27.0";


            var byteVersion = Encoding.ASCII.GetBytes(string.Concat("content=", "report_id=http%3A%2F%2Fpmp.acrowire.com%2Fpublic%2Findex.php%3Fpath_info%3Dtime%252F8&report%5Bcompany_id%5D=&report%5Buser_filter%5D=logged_user&report%5Bdate_filter%5D=all&report%5Bbillable_filter%5D=all&summarize%5Bsum_by_1%5D=&show_time_records=1&project_id=8&export_format=on_screen&is_export=0&create_invoice=0&submitted=submitted"));
            request.ContentLength = byteVersion.Length;
            request.Headers.Add("Cookie", cookie);

            var stream = request.GetRequestStream();
            stream.Write(byteVersion, 0, byteVersion.Length);
            stream.Close();

            var response = (HttpWebResponse)request.GetResponse();
            var responceStream = response.GetResponseStream();
            if (responceStream == null)
            {
                throw new Exception("Stream is null");
            }
            var oStreamReader = new StreamReader(responceStream, Encoding.UTF8);
            return oStreamReader.ReadToEnd();
        }
    }
}