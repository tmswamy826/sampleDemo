using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using System.Device.Location;

namespace TestConsole
{
    class WordDocReadText
    {
       void Main()
        {
            // Open a doc file.
            // Application application = new Application();
            // Document document = application.Documents.Open("C:\\word.doc");

            // Loop through all words in the document.
            int count=1;//= document.Words.Count;
            for (int i = 1; i <= count; i++)
            {
                // Write the word.
                string text = null; // = document.Words[i].Text;
                Console.WriteLine("Word {0} = {1}", i, text);
            }
            // Close word.
          //  application.Quit();
        }

       public void text( string searchTerm)
        {
            string filePath = @"C:\Users\52045046\Desktop\K.T.txt";
            System.IO.FileStream fs = System.IO.File.OpenRead(filePath);
            byte[] binaryData = new byte[fs.Length];
            fs.Read(binaryData, 0, binaryData.Length);
            string text = System.Text.Encoding.ASCII.GetString(binaryData);
            string[] paragraphs = text.Split('\r').Select(m => m.Trim('\n')).ToArray();

            for (int i=0;i<paragraphs.Length;i++)
            {
                int count = 0;
                int j = 0;
                while ((j= paragraphs[i].ToLower().IndexOf(searchTerm.ToLower(), j)) != -1)
                {
                    j += searchTerm.Length;
                    count++;
                }
                Console.WriteLine("{0} occurrences(s) of the search term \"{1}\" were found.", count, searchTerm);
            }

            Console.Read();
        }

        private double distance(double lat1, double lon1, double lat2, double lon2, char unit)
        {
            double theta = lon1 - lon2;
            double dist = Math.Sin(deg2rad(lat1)) * Math.Sin(deg2rad(lat2)) + Math.Cos(deg2rad(lat1)) * Math.Cos(deg2rad(lat2)) * Math.Cos(deg2rad(theta));
            dist = Math.Acos(dist);
            dist = rad2deg(dist);
            dist = dist * 60 * 1.1515;
            if (unit == 'K')
            {
                dist = dist * 1.609344;
            }
            else if (unit == 'N')
            {
                dist = dist * 0.8684;
            }
            return (dist);
        }

        //:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        //::  This function converts decimal degrees to radians             :::
        //:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        private double deg2rad(double deg)
        {
            return (deg * Math.PI / 180.0);
        }

        //:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        //::  This function converts radians to decimal degrees             :::
        //:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        private double rad2deg(double rad)
        {
            return (rad / Math.PI * 180.0);
        }

       // Console.WriteLine(distance(32.9697, -96.80322, 29.46786, -98.53506, "M"));
       // Console.WriteLine(distance(32.9697, -96.80322, 29.46786, -98.53506, "K"));
       // Console.WriteLine(distance(32.9697, -96.80322, 29.46786, -98.53506, "N"));
        public void FinalDistatce()
        {
            Mytest();
           // Console.WriteLine(distance(17.440082, 78.348917, 17.447412, 78.37623, 'K'));

            /* New Logic*/

            string sourceAddress = getAddress(17.4130944, 78.4662206); //(Lat1Source,Lan1Desti)
            string destinationAddress = getAddress(17.448294, 78.391487); //(lat2Source,Long2Dest)

            getDistance(sourceAddress, destinationAddress);
        }

        #region GPS Location
        protected string GetJsonData(string url)
        {
            string sContents = string.Empty;
            string me = string.Empty;
            try
            {
                if (url.ToLower().IndexOf("https:") > -1)
                {
                    System.Net.WebClient client = new System.Net.WebClient();
                    byte[] response = client.DownloadData(url);
                    sContents = System.Text.Encoding.ASCII.GetString(response);
                }
                else
                {
                    System.IO.StreamReader sr = new System.IO.StreamReader(url);
                    sContents = sr.ReadToEnd();
                    sr.Close();
                }
            }
            catch
            {
                sContents = "unable to connect to server ";
            }
            return sContents;
        }

        public string getAddress(double latitude, double longitude)
        {
            string address = "";
            string content = GetJsonData("https://maps.googleapis.com/maps/api/geocode/json?latlng=" + latitude + "," + longitude + "&sensor=true");
            JObject obj = JObject.Parse(content);
            try
            {
                address = obj.SelectToken("results[0].address_components[3].long_name").ToString();
                return address;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return address;
        }

        public string getDistance(string source, string destination)
        {
            int distance = 0;
            string content = GetJsonData("https://maps.googleapis.com/maps/api/directions/json?origin=" + source + "&destination=" + destination + "&sensor=false");
            JObject obj = JObject.Parse(content);
            try
            {
                distance = (int)obj.SelectToken("routes[0].legs[0].distance.value");
                return (distance / 1000).ToString() + " K.M.";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return (distance / 1000).ToString() + " K.M.";
        }

        #endregion


        // The coordinate watcher.
        private GeoCoordinateWatcher Watcher = null;

        // Create and start the watcher.
       // private void Form1_Load(object sender, EventArgs e)
       private void Location()
        {
            // Create the watcher.
            Watcher = new GeoCoordinateWatcher();

            // Catch the StatusChanged event.
           // Watcher.StatusChanged += Watcher_StatusChanged;

            // Start the watcher.
            Watcher.Start();
        }

        // The watcher's status has change. See if it is ready.
        private void Watcher_StatusChanged()
        {
            GeoCoordinateWatcher watcher1;
            watcher1 = new GeoCoordinateWatcher();

            watcher1.PositionChanged += (sender, e) =>
            {
                var coordinate = e.Position.Location;

                Console.WriteLine("Lat: {0}, Long: {1}", coordinate.Latitude,
                    coordinate.Longitude);
                
                // Uncomment to get only one event.
               // watcher1.Stop(); 

                //var lat = coordinate.Latitude;
                //var lon = coordinate.Longitude;
            };

            //Console.ReadLine();

            // Begin listening for location updates.
            watcher1.Start();

        }

        private void Mytest()
        {
            GeoCoordinateWatcher Wat = null;
            Wat = new GeoCoordinateWatcher();

            // Catch the StatusChanged event.
            Wat.PositionChanged += (sender, e) =>
            {
                var coordinate1 = e.Position.Location;

                Console.WriteLine("OUr Current Location " +"Lat: {0}, Long: {1}", coordinate1.Latitude,
                    coordinate1.Longitude);

                // Uncomment to get only one event.
                // watcher1.Stop(); 

                //var lat = coordinate.Latitude;
                //var lon = coordinate.Longitude;
            };
            // Start the watcher.
            Wat.Start();
        }
       
    }
}
