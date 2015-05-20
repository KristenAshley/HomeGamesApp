using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Configuration;

namespace Weather
{
    class WeatherProgram
    {

        public List<WundergroundForecastData> DasWetter;

        public struct WundergroundForecastData
        {

            public DateTime Date;
            public string high;
            public string low;
            public string conditions;
            public string avehumidity;
            public string avewind;
        }

        public WeatherProgram()
        {
            string wunderground_key = ConfigurationManager.AppSettings["APIKey"].ToString();
            DasWetter = parseForecast("http://api.wunderground.com/api/" + wunderground_key + "/forecast/q/" + ConfigurationManager.AppSettings["StateAbbreviation"].ToString() + "/" + ConfigurationManager.AppSettings["City"].ToString() + ".xml");
        }


        private static List<WundergroundForecastData> parseForecast(string input_xml)
        {

            var cli = new WebClient();
            string weather = cli.DownloadString(input_xml);
            string xmlWunderground = null;
            try
            {
                using (StreamReader sr = new StreamReader(cli.OpenRead(input_xml)))
                {
                    xmlWunderground = sr.ReadToEnd();
                    sr.Close();
                }
            }
            catch (Exception)
            {

            }

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlWunderground);
            XmlElement root = doc.DocumentElement;

            String nodelist = doc.GetElementsByTagName("simpleforecast").Item(0).InnerXml.ToString();

            doc = new XmlDocument();
            doc.LoadXml(nodelist);
            XmlNodeList nodelistChild = doc.GetElementsByTagName("forecastday");

            List<WundergroundForecastData> list = new List<WundergroundForecastData>();



            foreach (XmlNode a in nodelistChild)
            {

                WundergroundForecastData b = new WundergroundForecastData();
                XmlNode hightemp = a["high"];
                XmlNode lowtemp = a["low"];
                XmlNode date = a["date"];
                XmlNode conditions = a["conditions"];
                XmlNode avehumidity = a["avehumidity"];
                XmlNode avewind = a["avewind"];


                b.Date = new DateTime(Convert.ToInt32(date["year"].InnerText), Convert.ToInt32(date["month"].InnerText), Convert.ToInt32(date["day"].InnerText));
                b.high = hightemp["fahrenheit"].InnerXml.ToString();
                b.low = lowtemp["fahrenheit"].InnerXml.ToString();
                b.conditions = conditions.InnerXml.ToString();
                b.avehumidity = avehumidity.InnerXml.ToString();
                b.avewind = avewind["mph"].InnerText;

                list.Add(b);

            }

            return list;

        }
    }
}