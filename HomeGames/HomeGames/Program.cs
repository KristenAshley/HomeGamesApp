using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Data.Objects;
using System.IO;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Xml;
using System.Net;
using Weather;
using System.Configuration;
 
namespace HomeGames
{
    class Program
    {
        static void QueryCsv(WeatherProgram weather)
        {
            string filename = ConfigurationManager.AppSettings["CSVLocation"].ToString();
            OleDbConnection cn = new OleDbConnection(string.Format(@"Provider=Microsoft.Jet.OleDb.4.0; Data Source={0};Extended Properties=""Text;HDR=YES;FMT=Delimited""", Path.GetDirectoryName(filename)));
            OleDbCommand cmd = new OleDbCommand(@"SELECT * FROM  [" + Path.GetFileName(filename) + "]", cn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);

            cn.Open();

            DataTable dt = new DataTable();
            da.Fill(dt);

            DateTime date = DateTime.Today,
                start = date.Date.AddDays(-(int)date.DayOfWeek), // prev sunday 00:00
                end = start.AddDays(7); // next sunday 00:00

            //Upcoming Games for the Week
            var WeekGames = from r in dt.AsEnumerable()
                            where r.Field<string>(4) == ConfigurationManager.AppSettings["ParkName"].ToString() && r.Field<DateTime>(0) >= start && r.Field<DateTime>(0) < end
                            select new
                            {
                                startdate = r.Field<DateTime>(0).ToShortDateString(),
                                starttime = r.Field<string>(1),
                                subject = r.Field<string>(3),
                                dayoftheweek = r.Field<DateTime>(0).DayOfWeek
                            };


            //Daily Game Reminder or Weekend Game Reminder (if Friday)
            var DailyReminder = from d in dt.AsEnumerable()
                                where d.Field<string>(4) == ConfigurationManager.AppSettings["ParkName"].ToString()
                                &&
                                (
                                (
                                d.Field<DateTime>(0).ToShortDateString() == DateTime.Now.ToShortDateString()//DateTime.Now.AddDays(8).ToShortDateString()//DateTime.Now.ToShortDateString()
                                && DateTime.Now.DayOfWeek != DayOfWeek.Sunday// DateTime.Now.AddDays(8).DayOfWeek// DateTime.Now.DayOfWeek
                                && DateTime.Now.DayOfWeek != DayOfWeek.Saturday// DateTime.Now.AddDays(8).DayOfWeek// DateTime.Now.DayOfWeek
                                && DateTime.Now.DayOfWeek != DayOfWeek.Friday// DateTime.Now.AddDays(8).DayOfWeek// DateTime.Now.DayOfWeek
                                )
                                ||
                                (
                                DateTime.Now.DayOfWeek == DayOfWeek.Friday// DateTime.Now.AddDays(8).DayOfWeek// DateTime.Now.DayOfWeek
                                &&
                                (
                                d.Field<DateTime>(0).ToShortDateString() == DateTime.Now.ToShortDateString() //DateTime.Now.AddDays(8).ToShortDateString()//DateTime.Now.ToShortDateString()
                                || d.Field<DateTime>(0).ToShortDateString() == DateTime.Now.AddDays(1).ToShortDateString()
                                || d.Field<DateTime>(0).ToShortDateString() == DateTime.Now.AddDays(2).ToShortDateString()
                                )
                                )
                                )
                                select new
                                {
                                    startdate = d.Field<DateTime>(0).ToShortDateString(),
                                    starttime = d.Field<string>(1),
                                    subject = d.Field<string>(3),
                                    dayoftheweek = d.Field<DateTime>(0).DayOfWeek
                                };

            string emailBody = "";
            foreach (var game in WeekGames)
            {
                emailBody += (String.Format("{0} {1} {2} {3}", game.dayoftheweek, game.startdate, game.starttime, game.subject)) + Environment.NewLine;
            }

            string emailBodyDaily = "";
            bool emailcheck;
            foreach (var game in DailyReminder)
            {
                emailcheck = false;
                foreach (Weather.WeatherProgram.WundergroundForecastData day in weather.DasWetter)
                {
                    if ((day.Date.ToShortDateString() == game.startdate) && emailcheck == false)
                    {
                        emailBodyDaily += (String.Format("{0} {1} {2} \r\n{3} \r\nConditions:  {4} \r\nTemperature:  high- {5} °F low- {6} °F\r\nHumidity:  {7} %\r\nWind:  {8} mph\r\n", game.dayoftheweek, game.startdate, game.starttime, game.subject, day.conditions, day.high, day.low, day.avehumidity, day.avewind)) + Environment.NewLine;
                        emailcheck = true;
                        break;
                    }

                }
                if (emailcheck == false)
                {
                    emailBodyDaily += (String.Format("{0} {1} {2} \r\n{3} \r\n", game.dayoftheweek, game.startdate, game.starttime, game.subject)) + Environment.NewLine;
                }
            }

            if (DateTime.Now.DayOfWeek == DayOfWeek.Monday)
            {
                email WeeklyReminder = new email();
                WeeklyReminder.CreateEmailItem("This Week at " + ConfigurationManager.AppSettings["ParkName"].ToString(), ConfigurationManager.AppSettings["toEmail"].ToString(), emailBody);
            }

            if (emailBodyDaily != string.Empty)
            {
                email upcominggames = new email();
                upcominggames.CreateEmailItem("Upcoming Game(s)", ConfigurationManager.AppSettings["toEmail"].ToString(), emailBodyDaily);
            }
        }

        class email
        {

            public void CreateEmailItem(string subjectEmail, string toEmail, string bodyEmail)
            {
                try
                {
                    try
                    {
                        Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
                        Outlook.MailItem mailItem = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                        mailItem.Subject = subjectEmail;
                        mailItem.To = toEmail;
                        mailItem.Body = bodyEmail;
                        mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                        ((Outlook._MailItem)mailItem).Send();
                    }
                    catch
                    {
                        alternateEmail e = new alternateEmail(ConfigurationManager.AppSettings["username"].ToString(), ConfigurationManager.AppSettings["password"].ToString());
                        MailMessage msg = new MailMessage(ConfigurationManager.AppSettings["toEmail"].ToString(), ConfigurationManager.AppSettings["toEmail"].ToString());
                        msg.Subject = subjectEmail;
                        msg.Body = bodyEmail;
                        e.Send(msg);
                    }
                }
                catch
                {

                }

            }
        }

        static void Main(string[] args)
        {
            WeatherProgram a = new WeatherProgram();
            HomeGames.Program.QueryCsv(a);
        }
    }

}
