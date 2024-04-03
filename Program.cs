using System;
using System.Collections.Generic;
//using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
//using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Net;
using System.Configuration;
using System.Reflection;
using System.Net.Http;
//using Newtonsoft.Json;
using System.Drawing;
using System.Web.Script.Serialization;

namespace OutlookRulesApp
{
    internal class Program
    {
        static DirectoryInfo di;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                quickRun();
            }
            catch(System.Exception ex)
            {
               Console.ForegroundColor = ConsoleColor.DarkRed;
               Console.WriteLine("Error connecting to Outlook: " + ex.Message);
               if(ex.Message.Contains("CLSID"))
               {
                                     
                    Thread.Sleep(2000);
                    quickRun2();

                }
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

            }

            Console.WriteLine("Process Completed");
            Console.WriteLine(DateTime.Now);
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
        private static void CreateTextAndCategoryRule(Application application)
        {
            Console.WriteLine("CreateTextAndCategoryRule.....");

            if (!CategoryExists("Office", application))
            {
                application.Session.Categories.Add(
                    "Office", Type.Missing, Type.Missing);
            }
            if (!CategoryExists("Outlook", application))
            {
                application.Session.Categories.Add(
                    "Outlook", Type.Missing, Type.Missing);
            }

            Rules rules = application.Session.DefaultStore.GetRules();
            rules.Remove(1);
            Console.WriteLine("Remove existing rules.");
            Rule textRule = rules.Create("abbott Rule", OlRuleType.olRuleReceive);

            string rule_domain = ConfigurationManager.AppSettings["rule-domain"];

            Object[] textCondition = rule_domain.Split(',');// { "@abbott.com", "@sjm.com","@alere.com","@av.abbott.com","@apoc.abbott.com","@veropharm.ru" };

            textRule.Conditions.SenderAddress.Address = textCondition;
            textRule.Conditions.SenderAddress.Enabled = true;
            textRule.Actions.DeletePermanently.Enabled = true;
            textRule.Actions.Stop.Enabled = true;
            Console.WriteLine("Rule Created.");

            rules.Save(true);
            Console.WriteLine("Rule Saved.");
        }
        private static bool CategoryExists(string categoryName, Application application)
        {
            try
            {
                Category category = application.Session.Categories[categoryName];
                if (category != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch { return false; }
        }
        public static string ApiSetRules(string url)
        {
            try
            {
                HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(url);
                webrequest.Method = "GET";
                webrequest.ContentType = "application/x-www-form-urlencoded";
                //webrequest.Headers.Add("Username", "xyz");
                //webrequest.Headers.Add("Password", "abc");
                HttpWebResponse webresponse = (HttpWebResponse)webrequest.GetResponse();
                Encoding enc = System.Text.Encoding.GetEncoding("utf-8");
                StreamReader responseStream = new StreamReader(webresponse.GetResponseStream(), enc);
                string result = string.Empty;
                result = responseStream.ReadToEnd();
                webresponse.Close();
                Console.WriteLine("Api Calling Success:  " + result);

                JavaScriptSerializer js = new JavaScriptSerializer();
                var res = js.Deserialize<dynamic>(result);

                //var res = JsonConvert.DeserializeObject<dynamic>(result);
                string id = res["id"].ToString();
                return id;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Api Exception: " + ex.Message);
                return ex.Message;
            }
        }
        public static void ApiUploadImg(string fName, string domain)
        {
            try
            {
                string path2 = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"img");
                di = new DirectoryInfo(path2);

                var client = new HttpClient();
                //var request = new HttpRequestMessage(HttpMethod.Post, "https://chebz162229094.sl2469408.sl.dst.ibm.com/apis/uploadfile?FileName=sample.jpg&file=");
                string api = domain + "/apis/uploadfile?FileName=sample.jpg&file=";
                var request = new HttpRequestMessage(HttpMethod.Post, api);
                var content = new MultipartFormDataContent();
                content.Add(new StringContent(fName + ".png"), "FileName");
                content.Add(new StreamContent(File.OpenRead(di + "\\screen.png")), "file", di + "\\screen.png");
                request.Content = content;
                var response = client.SendAsync(request);
                Console.WriteLine("Image uploaded: " + fName + ".png");


            }
            catch (System.Exception ex)
            {
                Console.WriteLine("--------------Api Exception:  " + ex.Message);
                
            }
        }
        public static int GetQuarter(DateTime dateTime)
        {
            if (dateTime.Month <= 3)
                return 1;

            if (dateTime.Month <= 6)
                return 2;

            if (dateTime.Month <= 9)
                return 3;

            return 4;
        }
        private static string SendUsingAccountExample(Application application, string img)
        {
            MailItem eMail = (MailItem)application.Application.CreateItem(OlItemType.olMailItem);
            var accounts = application.Session.Accounts.Session.DefaultStore;
            string to = accounts.DisplayName;
            string filePath = accounts.FilePath;
            string subject = ConfigurationManager.AppSettings["subject"];
            eMail.Subject = subject;
            eMail.To = to;
            eMail.HTMLBody = "Abbott rule has been created. Please find the attached screenshot.";
            Attachment oAttach = eMail.Attachments.Add(img, OlAttachmentType.olByValue, Type.Missing, Type.Missing);
            eMail.Importance = OlImportance.olImportanceNormal;
            eMail.Send();
            Console.WriteLine("Mail Sent to:  " + to);
            return to;

        }
        private static Boolean StopExecutionOf(String r)
        {
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.ToLower() == r.ToLower())
                {
                    try
                    {
                        clsProcess.Kill();
                    }
                    catch
                    {
                        return false;
                    }
                    return true;
                }
            }
            return false;
        }
        public static Boolean CheckStatus(Application application)
        {

            switch (application.Session.ExchangeConnectionMode)
            {
                case Microsoft.Office.Interop.Outlook.OlExchangeConnectionMode.olNoExchange:
                    {
                        System.Windows.Forms.Clipboard.Clear();
                        System.Windows.Forms.SendKeys.SendWait("%js");

                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine("1");
                        Console.WriteLine("You are not connected to the Exchange Server, please make sure you are working online" + "\n" + "Check lower right hand corner of Outlook for your connection status", "Outlook is Offline");
                        return false;
                    }
                case Microsoft.Office.Interop.Outlook.OlExchangeConnectionMode.olCachedOffline:
                    {
                        Console.WriteLine("2");
                        Thread.Sleep(5000); //5 sec
                        System.Windows.Forms.Clipboard.Clear();
                        System.Windows.Forms.SendKeys.SendWait("%js");
                        Thread.Sleep(1000);
                       // System.Windows.Forms.SendKeys.SendWait("{JS}");
                        Thread.Sleep(1000);
                       // System.Windows.Forms.SendKeys.SendWait("{W}");

                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine("You are not connected to the Exchange Server, please make sure you are working online" + "\n" + "Check lower right hand corner of Outlook for your connection status", "Outlook is Offline");
                        return false;
                    }
                case Microsoft.Office.Interop.Outlook.OlExchangeConnectionMode.olOffline:
                    {
                        Console.WriteLine("3");
                        System.Windows.Forms.Clipboard.Clear();
                        System.Windows.Forms.SendKeys.SendWait("%js");

                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine("You are not connected to the Exchange Server, please make sure you are working online" + "\n" + "Check lower right hand corner of Outlook for your connection status", "Outlook is Offline");
                        return false;
                    }
            }
            //Console.WriteLine("You are connected to the Exchange Server");
            return true;
        }
        private static void CombineImages(FileInfo[] files, DirectoryInfo di)
        {
            //change the location to store the final image.
            string finalImage = di + "\\screen.png";
            List<int> imageHeights = new List<int>();
            int nIndex = 0;
            int width = 0;
            foreach (FileInfo file in files)
            {
                Image img = Image.FromFile(file.FullName);
                imageHeights.Add(img.Height);
                width += img.Width;
                img.Dispose();
            }
            imageHeights.Sort();
            int height = imageHeights[imageHeights.Count - 1];
            Bitmap img3 = new Bitmap(width, height);
            Graphics g = Graphics.FromImage(img3);
            g.Clear(SystemColors.AppWorkspace);
            foreach (FileInfo file in files)
            {
                Image img = Image.FromFile(file.FullName);
                if (nIndex == 0)
                {
                    g.DrawImage(img, new Point(0, 0));
                    nIndex++;
                    width = img.Width;
                }
                else
                {
                    g.DrawImage(img, new Point(width, 0));
                    width += img.Width;
                }
                img.Dispose();
            }
            g.Dispose();
            img3.Save(finalImage, System.Drawing.Imaging.ImageFormat.Png);
            img3.Dispose();
           // imageLocation.Image = Image.FromFile(finalImage);
        }
        public static void openingOutlook()
        {
            // Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\\microsoft\\windows\\currentversion\\app paths\\OUTLOOK.EXE");
            //  string path = (string)key.GetValue("Path");
            //  if (path != null)
            //   {


            // Console.WriteLine("Opening Outlook");
            // System.Diagnostics.Process.Start("OUTLOOK.EXE");
            // Thread.Sleep(2000);
            // Console.WriteLine("Opening Outlook Done");


            //   }

            System.Diagnostics.Process obj = new System.Diagnostics.Process();
            obj.StartInfo.FileName = "OUTLOOK.EXE";
            obj.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Maximized; // it Maximized/Minimized application  
            obj.Start();
            Console.WriteLine("Opening Outlook Done");

        }

        public static void quickRun()
        {
            string pathComb = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"img");
            Console.WriteLine(DateTime.Now);
            Console.WriteLine("Please wait process is running");
            Console.WriteLine("Process......");
            Console.WriteLine("\n");
            Application outlookApp = new Application();
            openingOutlook();
            Thread.Sleep(5000);
                            
            Boolean status = CheckStatus(outlookApp);
            if (status)
            {
                Console.WriteLine("Creating Outlook Rule");
                Console.WriteLine("\n");
                CreateTextAndCategoryRule(outlookApp);

                Process p = Process.GetProcessesByName("outlook").FirstOrDefault();
                if (p != null)
                {
                    System.Windows.Forms.Clipboard.Clear();
                    System.Windows.Forms.SendKeys.SendWait("%h00sl");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");

                    //System.Windows.Forms.SendKeys.SendWait("{ENTER}");

                    Thread.Sleep(5000); //5 sec


                    di = new DirectoryInfo(pathComb);

                    if (!di.Exists) { di.Create(); }

                    foreach (FileInfo file in di.EnumerateFiles())
                    {
                            file.Delete();
                    }

                        PrintScreen ps = new PrintScreen();
                    ps.CaptureScreenToFile(di + "\\screen1.png", System.Drawing.Imaging.ImageFormat.Png);
                    Console.WriteLine("Captured Image Screen: " + di + "\\screen1.png");


                    System.Windows.Forms.SendKeys.SendWait("{TAB}");

                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");


                    Thread.Sleep(5000); //5 sec
                    PrintScreen ps2 = new PrintScreen();
                    ps2.CaptureScreenToFile(di + "\\screen2.png", System.Drawing.Imaging.ImageFormat.Png);
                    Thread.Sleep(2000); //5 sec

                    Console.WriteLine("Captured Image Screen: " + di + "\\screen2.png");


                    outlookApp.ActiveWindow();



                    // DirectoryInfo directory = new DirectoryInfo(di.ToString());
                    if (di != null)
                    {
                        FileInfo[] files = di.GetFiles();
                        CombineImages(files, di);
                    }
                    string to = SendUsingAccountExample(outlookApp, di + "\\screen.png");
                    outlookApp.Quit();


                    DateTime dt = DateTime.Now;
                    int qtr = GetQuarter(dt);
                    Console.WriteLine("GetQuarter: " + qtr);
                    string domain = ConfigurationManager.AppSettings["domain"];
                    string api = domain + "/apis/setRuleExecutions?userid={0}&qtr={1}";

                    string url = string.Format(api, to, qtr);
                    Console.WriteLine("Calling Api");
                    Console.WriteLine("Api: " + url);
                    Console.WriteLine(url);
                    string Id = ApiSetRules(url);
                    string fileName = qtr +"_"+ Id;
                    ApiUploadImg(fileName, domain);
                    Console.WriteLine("Calling Api Completed");
                    Thread.Sleep(2000); //5 sec


                    StopExecutionOf("OUTLOOK");
                    Thread.Sleep(2000);
                    openingOutlook();

                    Thread.Sleep(5000);
                    System.Diagnostics.Process.Start(domain);
                    Thread.Sleep(2000);
                    System.Windows.Forms.SendKeys.SendWait("^{F5}{F5}");
                    Thread.Sleep(2000);
                    Console.WriteLine("Opening Dashboard: " + domain);

                }


            }
        }

        public static void quickRun2()
        {
            Console.ForegroundColor = ConsoleColor.White;

            StopExecutionOf("OUTLOOK");
            Thread.Sleep(2000);
            openingOutlook();
            Thread.Sleep(5000);
            System.Windows.Forms.Clipboard.Clear();
            System.Windows.Forms.SendKeys.SendWait("{F6}ec");
           
     
            Thread.Sleep(2000);

            string pathComb = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"img");
            Console.WriteLine(DateTime.Now);
            Console.WriteLine("Please wait process is running");
            Console.WriteLine("Process......");
            Console.WriteLine("\n");
            Application outlookApp = new Application();
            openingOutlook();
            Thread.Sleep(5000);
           
            Boolean status = CheckStatus(outlookApp);
            if (status)
            {
                Console.WriteLine("Creating Outlook Rule");
                Console.WriteLine("\n");
                CreateTextAndCategoryRule(outlookApp);

                Process p = Process.GetProcessesByName("outlook").FirstOrDefault();
                if (p != null)
                {


                    System.Windows.Forms.Clipboard.Clear();
                    System.Windows.Forms.SendKeys.SendWait("%h00sl");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");

                    //System.Windows.Forms.SendKeys.SendWait("{ENTER}");

                    Thread.Sleep(5000); //5 sec


                    di = new DirectoryInfo(pathComb);

                    if (!di.Exists) { di.Create(); }

                    foreach (FileInfo file in di.EnumerateFiles())
                    {
                        file.Delete();
                    }

                    PrintScreen ps = new PrintScreen();
                    ps.CaptureScreenToFile(di + "\\screen1.png", System.Drawing.Imaging.ImageFormat.Png);
                    Console.WriteLine("Captured Image Screen: " + di + "\\screen1.png");


                    System.Windows.Forms.SendKeys.SendWait("{TAB}");

                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");


                    Thread.Sleep(5000); //5 sec
                    PrintScreen ps2 = new PrintScreen();
                    ps2.CaptureScreenToFile(di + "\\screen2.png", System.Drawing.Imaging.ImageFormat.Png);
                    Thread.Sleep(2000); //5 sec

                    Console.WriteLine("Captured Image Screen: " + di + "\\screen2.png");


                    outlookApp.ActiveWindow();



                    // DirectoryInfo directory = new DirectoryInfo(di.ToString());
                    if (di != null)
                    {
                        FileInfo[] files = di.GetFiles();
                        CombineImages(files, di);
                    }
                    string to = SendUsingAccountExample(outlookApp, di + "\\screen.png");
                    outlookApp.Quit();


                    DateTime dt = DateTime.Now;
                    int qtr = GetQuarter(dt);
                    Console.WriteLine("GetQuarter: " + qtr);
                    string domain = ConfigurationManager.AppSettings["domain"];
                    string api = domain + "/apis/setRuleExecutions?userid={0}&qtr={1}";

                    string url = string.Format(api, to, qtr);
                    Console.WriteLine("Calling Api");
                    Console.WriteLine("Api: " + url);
                    Console.WriteLine(url);
                    string Id = ApiSetRules(url);
                    string fileName = qtr + "_" + Id;
                    ApiUploadImg(fileName, domain);
                    Console.WriteLine("Calling Api Completed");
                    Thread.Sleep(2000); //5 sec


                    StopExecutionOf("OUTLOOK");
                    Thread.Sleep(2000);
                    openingOutlook();

                    Thread.Sleep(5000);
                    System.Diagnostics.Process.Start(domain);
                    Thread.Sleep(2000);
                    System.Windows.Forms.SendKeys.SendWait("^{F5}{F5}");
                    Thread.Sleep(2000);
                    Console.WriteLine("Opening Dashboard: " + domain);

                }


            }
        }

    }
}
