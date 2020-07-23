using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace CalendarQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string ApplicationName = "Google Calendar API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;
            System.IO.StreamWriter sw = new System.IO.StreamWriter("Calender.txt");

            using (var stream =  new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
               // string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                 null).Result;
               // new FileDataStore(credPath, true)).Result;
             //   Console.WriteLine("Credential file saved to: " + credPath);
            }

            

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = DateTime.Today.AddDays(1);
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 40;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events.
            Events events = request.Execute();



            // 計數器 只顯示三天行程

            int weekcount = 0;
            int n = 4;
            if (events.Items != null && events.Items.Count > 0)
            {
                int ordercount = 0;

                string Schedule_Date = "";
                string weekday = "";
                foreach (var eventItem in events.Items)

                {

                if (weekcount<n)
 
                {
                        string whenstart = eventItem.Start.DateTime.ToString();
                        string whenend = eventItem.End.DateTime.ToString();



                        string ordername = "";


                        // 檢查日期是否相同
                        if (Schedule_Date != whenstart.Substring(whenstart.IndexOf("/") + 1, whenstart.IndexOf(" ") - whenstart.IndexOf("/") - 1))
                        {

                            weekcount++;

                            // 計數器 顯示一、二、三於此歸零

                            ordercount = 0;

                            IFormatProvider culture = new System.Globalization.CultureInfo("zh-TW", true);
                            DateTime dt = DateTime.ParseExact(whenstart.Substring(0, whenstart.IndexOf(" ")), "yyyy/M/d", culture);

                            switch (dt.DayOfWeek.ToString())
                            {
                                case "Monday":
                                    weekday = "(一)";
                                    break;
                                case "Tuesday":
                                    weekday = "(二)";
                                    break;
                                case "Wednesday":
                                    weekday = "(三)";
                                    break;
                                case "Thursday":
                                    weekday = "(四)";
                                    break;
                                case "Friday":
                                    weekday = "(五)";
                                    break;
                                case "Saturday":
                                    weekday = "(六)";
                                    break;
                                case "Sunday":
                                    weekday = "(日)";
                                    break;
                            }


                            Schedule_Date = whenstart.Substring(whenstart.IndexOf("/") + 1, whenstart.IndexOf(" ") - whenstart.IndexOf("/") - 1);

                            if (weekcount < n)

                            {
                                sw.WriteLine("---------------------------------------------------------");
                                sw.WriteLine("{0}{1}行程：", Schedule_Date, weekday);
                            }
                        }



                        if (String.IsNullOrEmpty(whenstart))
                        {
                            whenstart = "尚未安排";

                            whenend = "";
                        }



                        //開始時間擷取
                        int ss = whenstart.IndexOf("午");
                        string s = whenstart.Substring(ss + 2, 2);
                        if (whenstart.Contains("下午"))
                        {

                            if (whenstart.Contains("下午 12"))
                            {
                                whenstart = s + whenstart.Substring(ss + 4, 3);
                            }
                            else
                            {
                                s = (int.Parse(s) + 12).ToString();
                                whenstart = s + whenstart.Substring(ss + 4, 3);
                            }
                        }
                        else
                        {
                            whenstart = s + whenstart.Substring(ss + 4, 3);
                        }
                        //結束時間擷取
                        int ee = whenend.IndexOf("午");
                        string e = whenend.Substring(ee + 2, 2);
                        if (whenend.Contains("下午"))
                        {

                            if (whenend.Contains("下午 12"))
                            {
                                whenend = e + whenend.Substring(ee + 4, 3);
                            }
                            else
                            {
                                e = (int.Parse(e) + 12).ToString();
                                whenend = e + whenend.Substring(ee + 4, 3);
                            }
                        }
                        else
                        {
                            whenend = e + whenend.Substring(ee + 4, 3);
                        }
                        // 計數器 顯示一、二、三
                        ordercount++;

                        switch (ordercount)
                        {
                            case 1:
                                ordername = "一、";
                                break;
                            case 2:
                                ordername = "二、";
                                break;
                            case 3:
                                ordername = "三、";
                                break;
                            case 4:
                                ordername = "四、";
                                break;
                            case 5:
                                ordername = "五、";
                                break;
                            case 6:
                                ordername = "六、";
                                break;
                            case 7:
                                ordername = "七、";
                                break;
                            case 8:
                                ordername = "八、";
                                break;
                            case 9:
                                ordername = "九、";
                                break;
                            case 10:
                                ordername = "十、";
                                break;
                            case 11:
                                ordername = "十一、";
                                break;
                            case 12:
                                ordername = "十二、";
                                break;
                            case 13:
                                ordername = "十三、";
                                break;
                            case 14:
                                ordername = "十四、";
                                break;
                            case 15:
                                ordername = "十五、";
                                break;
                        }


                        if (weekcount < n)

                        {
                            sw.WriteLine("{0}{1}-{2}", ordername, whenstart, whenend);
                            sw.WriteLine("{0}", eventItem.Summary);
                            sw.WriteLine("地點：{0}", eventItem.Location);
                            sw.WriteLine("");
                        }
                    }
               
                }
                sw.WriteLine("以上。");

            }
            //    9/10(一)行程：
            //    一、09:30-10:00
            //    局長致詞-櫻花達人講座研習座談會
            //    地點：會中喽01會議室

            //    二、10:30 - 11:00
            //    市長致詞 - 光明會
            //    地點：外國
            else
            {
                sw.WriteLine("未搜尋到任何資料");
            }

            //// Create a string array that consists of three lines.
            //string[] lines = { "First line2", "Second line2", "Third line2" };
    
            //// WriteAllLines creates a file, writes a collection of strings to the file,
            //// and then closes the file.  You do NOT need to call Flush() or Close().
            //System.IO.File.WriteAllLines(@"WriteLines.txt", lines);

            //// Example #2: Write one string to a text file.
            //string text = "A class is the most powerful data type in C#. Like a structure, " +
            //               "a class defines the data and behavior of the data type. ";
            //// WriteAllText creates a file, writes the specified string to the file,
            //// and then closes the file.    You do NOT need to call Flush() or Close().
            //System.IO.File.WriteAllText(@"C:\Users\Kerla\Desktop\WriteText.txt", text);


            Console.Read();
            sw.Close();
        }
    }
}