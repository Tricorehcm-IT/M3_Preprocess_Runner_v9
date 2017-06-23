using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace M3_Preprocess_Runner
{
    class PEClass
    {
        public delegate void PublishMessageDelegate(string Key, string MessageToPublish);       // Step 1 for event handling, declare a delegate
        public event PublishMessageDelegate MessageToPublishEvent;                              // Step 2 for event handling, declare an event
        // Step 3 for event handling, assign a handler to listen for this event (usually in another class)

        public PEClass()
        { 
        }

        public void LoginAndSync()
        {
            // To login to PayEntry and sync using the TSR account... ( Discovered using Fiddler, EP )
            //"https://www4.payentry.com/103/Login.asp?dsn=M103&sbid=1";
            //POST: "viewChanged=1&viewCommand=login&viewArg=&dsn=M103&sbid=1&target=%2F103%2FCompany%2FMain.asp&bypass=&overrideaccess=&username=TSR&password=Tricore2%23&co=";
            //"https://www4.payentry.com/103/System/Sync/UpstreamTab.asp";
            //POST: "viewChanged=0&viewCommand=sync&viewArg=&SSyncDbInfo_00347F5F_2DAC69_2D4EC2_2D8AB9_2DE9AFD18CCBBA_hostname=tricorenj.com&SSyncDbInfo_00347F5F_2DAC69_2D4EC2_2D8AB9_2DE9AFD18CCBBA_port=82&add_co="
            //Above the GUID identifies the Tricore database (SB103), in M3 on the System > Setup > Server Info tab
            
            string dataToPost;
            string targetURI;
            string rootURI;
            Uri url;
            UriBuilder urlBuilder;

            // Get Redirect
            targetURI = "https://www.payentry.com/103";

            url = new Uri(targetURI);
            HttpWebRequest request = null;

            ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);
            CookieContainer cookieJar = new CookieContainer();

            request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = cookieJar;
            request.Method = "GET";
            HttpStatusCode responseStatus;

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                responseStatus = response.StatusCode;
                url = request.Address;
                if (responseStatus == HttpStatusCode.OK)
                {
                    rootURI = response.ResponseUri.ToString();
                    //MessageBox.Show(rootURI);
                }
                else
                {
                    //MessageBox.Show("Redirect failed");
                    return;
                }
            }

            // Login
            this.MessageToPublishEvent("Status", "PayEntry login in-progress...");

            targetURI = rootURI.Substring(0, rootURI.IndexOf(".payentry.com/103")) + ".payentry.com/103/Login.asp?dsn=M103&sbid=1"; //"https://www4.payentry.com/103/Login.asp?dsn=M103&sbid=1";
            dataToPost = "viewChanged=1&viewCommand=login&viewArg=&dsn=M103&sbid=1&target=%2F103%2FCompany%2FMain.asp&bypass=&overrideaccess=&username=TSR&password=Tricore2%23&co=";
            //PostToURL(dataToPost, targetURI);

            url = new Uri(targetURI);
            request = null;

            ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);
            cookieJar = new CookieContainer();

            request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = cookieJar;
            request.Method = "GET";

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                responseStatus = response.StatusCode;
                url = request.Address;
            }

            if (responseStatus == HttpStatusCode.OK)
            {
                urlBuilder = new UriBuilder(url);
                //urlBuilder.Path = urlBuilder.Path.Remove(urlBuilder.Path.LastIndexOf('/')); // +"/j_security_check"; //not necessary

                request = (HttpWebRequest)WebRequest.Create(urlBuilder.ToString());
                request.Referer = url.ToString();
                request.CookieContainer = cookieJar;
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";

                using (Stream requestStream = request.GetRequestStream())
                using (StreamWriter requestWriter = new StreamWriter(requestStream, Encoding.ASCII))
                {
                    string postData = dataToPost;
                    requestWriter.Write(postData);
                }
                string responseContent = null;
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (Stream responseStream = response.GetResponseStream())
                using (StreamReader responseReader = new StreamReader(responseStream))
                {
                    responseContent = responseReader.ReadToEnd();
                }
                //Console.WriteLine(responseContent);
                this.MessageToPublishEvent("Status", "PayEntry login complete");
            }
            else
            {
                //Console.WriteLine("Client was unable to connect!");
                this.MessageToPublishEvent("Status", "PayEntry login failed!");
                return;
            }

            // Sync
            targetURI = rootURI.Substring(0, rootURI.IndexOf(".payentry.com/103")) + ".payentry.com/103/System/Sync/UpstreamTab.asp";  // "https://www4.payentry.com/103/System/Sync/UpstreamTab.asp";
            dataToPost = "viewChanged=0&viewCommand=sync&viewArg=&SSyncDbInfo_00347F5F_2DAC69_2D4EC2_2D8AB9_2DE9AFD18CCBBA_hostname=tricorenj.com&SSyncDbInfo_00347F5F_2DAC69_2D4EC2_2D8AB9_2DE9AFD18CCBBA_port=82&add_co=";

            url = new Uri(targetURI);
            urlBuilder = new UriBuilder(url);

            request = (HttpWebRequest)WebRequest.Create(urlBuilder.ToString());
            request.Referer = url.ToString();
            request.CookieContainer = cookieJar;
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            
            using (Stream requestStream = request.GetRequestStream())
            using (StreamWriter requestWriter = new StreamWriter(requestStream, Encoding.ASCII))
            {
                string postData = dataToPost;
                requestWriter.Write(postData);
            }
            string responseContent2 = null;
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream responseStream = response.GetResponseStream())
            using (StreamReader responseReader = new StreamReader(responseStream))
            {
                responseContent2 = responseReader.ReadToEnd();
                responseStatus = response.StatusCode; 
            }
            if (responseStatus == HttpStatusCode.OK)
            {
                this.MessageToPublishEvent("Status", "Sync request submitted");
            }
            else
            {
                this.MessageToPublishEvent("Status", "Could not submit a sync request");
                return;
            }

            // Log-out
            targetURI = rootURI.Substring(0, rootURI.IndexOf(".payentry.com/103")) + ".payentry.com/103/Logout.asp?dsn=M103&sbid=1"; // "https://www.payentry.com/103/Logout.asp?dsn=M103&sbid=1"; //log-off 

            url = new Uri(targetURI);
            request = null;

            ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);
            cookieJar = new CookieContainer();

            request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = cookieJar;
            request.Method = "GET";

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                responseStatus = response.StatusCode;
                url = request.Address;
            }

            if (responseStatus == HttpStatusCode.OK)
            {
                //MessageBox.Show("Logout complete.\r\n\r\n" + responseStatus.ToString());
                this.MessageToPublishEvent("Status", "PayEntry logout complete");
            }
            else
            {
                //MessageBox.Show("Logout failed.\r\n\r\n" + responseStatus.ToString());
                this.MessageToPublishEvent("Status", "PayEntry logout failed");
            }
        }
    }
}
