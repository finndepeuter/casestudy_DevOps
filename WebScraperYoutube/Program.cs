// import of packages
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.IO;
using System.Reflection;
using System.Threading;
using OpenQA.Selenium.Safari;
using System.Collections.Generic;
using System.Web;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using static System.Net.WebRequestMethods;
using File = System.IO.File;
using System.Security.Cryptography.X509Certificates;
using System.Diagnostics;
using System.Threading.Channels;
using static DevOpsCaseStudy.ScrapingTest;
using System.Xml.Serialization;

namespace DevOpsCaseStudy
{
    public class ScrapingTest
    {
        static string url = "example";
        static string url1 = "https://www.youtube.com";
        static string url2 = "http://ictjobs.be/nl/";
        static string url3 = "https://www.swimrankings.net/index.php?page=athleteSelect&nationId=0&selectPage=SEARCH";
        //static string url3 = "https://www.swimrankings.net/index.php?page=athleteDetail&athleteId=4777812";
        static Int32 vcount = 1;
        static Int32 jcount = 1;
        static Int32 scount = 0;
        public IWebDriver? driver;

        public static void Main()
        {
            // do-while loop that keeps repeating the tool so it can be used multiple times
            do
            {
                var ScrapingTest = new ScrapingTest();
                // choose program
                Choose_Program();

                // start browser
                ChromeOptions capabilities = new ChromeOptions();
                capabilities.BrowserVersion = "107.0";
                Dictionary<string, object> ltOptions = new Dictionary<string, object>();
                ltOptions.Add("platformName", "Windows 10");
                capabilities.AddAdditionalOption("LT:Options", ltOptions);

                ChromeDriver driver = new ChromeDriver(capabilities);

                if (url == url1)
                {
                    // Scrape Youtube
                    ScrapingTest.YoutubeScraping(driver);
                    ScrapingTest.Close_Browser(driver);
                }
                else if (url == url2)
                {
                    // Scrape ICT-jobs
                    ScrapingTest.ICTScraping(driver);
                    ScrapingTest.Close_Browser(driver);
                }
                else
                {
                    // Scrape Swimrankings
                    ScrapingTest.SwimrankingsScraping(driver);
                    ScrapingTest.Close_Browser(driver);
                }
            }
            // when the url has been set to nothing again by the choose_program the while loop is stopped here
            while (url != "");

        }

        static ChromeDriver Start_Browser()
        {
            /* Local Selenium WebDriver */
            ChromeOptions capabilities = new ChromeOptions();
            capabilities.BrowserVersion = "107.0";
            Dictionary<string, object> ltOptions = new Dictionary<string, object>();
            ltOptions.Add("platformName", "Windows 10");
            capabilities.AddAdditionalOption("LT:Options", ltOptions);

            ChromeDriver driver = new ChromeDriver(capabilities);
            return (driver);

        }

        static string Choose_Program()
        {
            Int32 choice = 0;

            while (choice != 1 && choice != 2 && choice != 3 && choice != 4) { 
                // ask which tool the user wants to use
                Console.WriteLine("Make your choice for the tool with the coresponding number.");
                Console.WriteLine("1. Youtube Scraper");
                Console.WriteLine("2. ICT-jobs");
                Console.WriteLine("3. Swimrankings");
                Console.WriteLine("4. Quit");
                Console.WriteLine();

                // get the choice the user makes and convert it into an integer
                choice = Convert.ToInt32(Console.ReadLine());
            }

            // change the url according to the choice made
            if (choice == 1)
            {
                url = url1;
            }
            else if (choice == 2)
            {
                url = url2;
            }
            else if (choice == 3)
            {
                url = url3;
            } 
            else if (choice == 4)
            {
                // stop the tool and reset the url, this is only for when you are running the tool another time
                url = "";
                System.Environment.Exit(0);
            }
            Console.WriteLine("You chose " + choice);
            Console.WriteLine();
            return (url);

        }

        public IWebDriver? GetDriver() => driver;

        // method for scraping Youtube
        public class youtubedata
        {
            public int Id { get; set; }
            public string Title { get; set; }
            public string Channel { get; set; }
            public string Views { get; set; }
            public string Link { get; set; }
        }

        public void YoutubeScraping(IWebDriver? driver)
        {
            driver.Url = url;

            // create a csv with headings
            String youtube_file = @"C:\Users\finnd\source\repos\files\youtube_output.csv";
            String seperator = ";";
            StringBuilder output = new StringBuilder();
            String[] headings = { "VideoID", "VideoTitle", "YoutubeChannel", "Views", "VideoLink" };
            output.AppendLine(string.Join(seperator, headings));

            // create list to add to jsonfile
            List<youtubedata> _youtubedata = new List<youtubedata>();

            /* Explicit Wait to ensure that the page is loaded completely by reading the DOM state */
            var timeout = 10000; /* Maximum wait time of 10 seconds */
            var wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(timeout));
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            Thread.Sleep(3000);

            // accept cookies
            IWebElement cookies = driver.FindElement(By.XPath(".//*[@id=\"content\"]/div[2]/div[6]/div[1]/ytd-button-renderer[2]/yt-button-shape/button/yt-touch-feedback-shape/div/div[2]"));
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
            executor.ExecuteScript("arguments[0].click();", cookies);

            /* implement search option */
            // Ask the user for a searchterm
            Console.WriteLine("What do you want to search for?");
            string inputUser = Console.ReadLine();

            // use the path to find the searchbar and insert the searchterm the user chose there
            var searchOption = driver.FindElement(By.XPath("/html/body/ytd-app/div/div[1]/ytd-masthead/div[3]/div[2]/ytd-searchbox/form/div[1]/div[1]/input"));
            searchOption.SendKeys(inputUser);
            searchOption.Submit();
            Console.WriteLine("Your search is processing...");
            Console.WriteLine();
            Thread.Sleep(1000);
            
            // select the filter option
            IWebElement filter = driver.FindElement(By.XPath(".//*[@id=\"container\"]/ytd-toggle-button-renderer/yt-button-shape/button"));
            IJavaScriptExecutor filterExecutor = (IJavaScriptExecutor)driver;
            filterExecutor.ExecuteScript("arguments[0].click();", filter);
            
            // click the most recently uploaded option to sort on
            IWebElement most_recent = driver.FindElement(By.XPath(".//*[@id=\"label\"]/yt-formatted-string"));
            IJavaScriptExecutor mostrecentExecutor = (IJavaScriptExecutor)driver;
            mostrecentExecutor.ExecuteScript("arguments[0].click();", most_recent);

            /* Explicit Wait to ensure that the page is loaded completely by reading the DOM state */
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            Thread.Sleep(3000);

            /* Select the part on the page where the videos are situated */
            // By elem_videos = By.CssSelector("#contents");
            By elem_videos = By.CssSelector("#contents > ytd-video-renderer");
            ReadOnlyCollection<IWebElement> videos = driver.FindElements(elem_videos);
            Console.WriteLine("The 5 most recently uploaded videos are:");

            /* Go through the Videos List and scrap the same to get the attributes of the videos*/
            foreach (IWebElement video in videos)
            {
                string str_title, str_views, str_uploader, str_link;
                // search for the title
                IWebElement elem_video_title = video.FindElement(By.CssSelector("#video-title"));
                str_title = elem_video_title.GetAttribute("title");
                // search for the views
                IWebElement elem_video_views = video.FindElement(By.XPath(".//*[@id=\"metadata-line\"]/span[1]"));
                str_views = elem_video_views.Text;
                // search for the channelname
                IWebElement elem_video_uploader = video.FindElement(By.XPath(".//*[@id=\"text\"]/a"));
                str_uploader = elem_video_uploader.GetAttribute("innerText");
                // search for the videolink
                IWebElement elem_video_link = video.FindElement(By.XPath(".//*[@id=\"thumbnail\"]"));
                str_link = elem_video_link.GetAttribute("href");

                Console.WriteLine("******* Video " + (vcount) + " *******");
                Console.WriteLine("Video Title: " + str_title);
                Console.WriteLine("Video Views: " + str_views);
                Console.WriteLine("Channel: " + str_uploader);
                Console.WriteLine("Link: " + str_link);
                Console.WriteLine("\n");

                // create newline to append to csvfile
                string newLine = string.Format("{0}, {1}, {2}, {3}, {4}",
                    vcount.ToString(), str_title, str_uploader, str_views, str_link);
                output.AppendLine(string.Join(seperator, newLine));

                // add new video to jsonfile
                _youtubedata.Add(new youtubedata()
                {
                    Id = vcount,
                    Title = str_title,
                    Channel = str_uploader,
                    Views = str_views,
                    Link = str_link
                });

                vcount+= 1;
                // we only need the first 5 videos so we use a break to stop after 5 videos
                if (vcount > 5) {
                    // write all lines to a jsonfile
                    string json = JsonSerializer.Serialize(_youtubedata);
                    File.WriteAllText(@"C:\Users\finnd\source\repos\files\youtube.json", json);
                    break;
                }
            }
            // write all lines to csv file and give error if something goes wrong
            try
            {
                File.AppendAllText(youtube_file, output.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine("Data could not be written to the CSV file.");
                return;
            }

            Console.WriteLine("Scraping Data from the 5 most recently uploaded Youtube videos done and writen to a file.");

        }

        // method for scraping ICTjobs
        public class ictdata
        {
            public int Id { get; set; }
            public string Title { get; set; }
            public string Company { get; set; }
            public string Location { get; set; }
            public string Keywords { get; set; }
            public string Link { get; set; }
        }

        public void ICTScraping(IWebDriver? driver)
        {
            driver.Url = url;

            // create a csv with headings
            String ict_file = @"C:\Users\finnd\source\repos\files\ict_output.csv";
            String separator = ";";
            StringBuilder output = new StringBuilder();
            String[] headings = { "JobID", "JobTitle", "Company", "Location", "Keywords", "JobLink" };
            output.AppendLine(string.Join(separator, headings));

            // create list to add to jsonfile
            List<ictdata> _ictdata = new List<ictdata>();

            /* Explicit Wait to ensure that the page is loaded completely by reading the DOM state */
            var timeout = 10000; /* Maximum wait time of 10 seconds */
            var wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(timeout));
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            Thread.Sleep(5000);

            // implement search option
            Console.WriteLine("What do you want to search for?");
            string inputUser = Console.ReadLine();

            var searchOption = driver.FindElement(By.XPath("//*[@id=\"keywords-input\"]"));
            searchOption.SendKeys(inputUser);
            searchOption.Submit();
            Console.WriteLine("Your search is processing...");

            /* The page only shows 1 job without scroling, to fix this we use this part 
             to scroll through the page far enough for 5 jobs to load*/
            Int64 last_height = (Int64)(((IJavaScriptExecutor)driver).ExecuteScript("return document.documentElement.scrollHeight"));
            while (true)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(0, document.documentElement.scrollHeight);");
                /* Wait to load page */
                Thread.Sleep(1000);
                /* Calculate new scroll height and compare with last scroll height */
                Int64 new_height = (Int64)((IJavaScriptExecutor)driver).ExecuteScript("return document.documentElement.scrollHeight");
                if (new_height > 2000)
                    /* If heights are the same it will exit the function */
                    break;
                last_height = new_height;
            }

            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            Thread.Sleep(3000);

            By elem_jobs = By.CssSelector("#search-result-body > div > ul > li.clearfix");
            ReadOnlyCollection<IWebElement> jobs = driver.FindElements(elem_jobs);
            Console.WriteLine("The 5 most recent jobs:");

            /* Go through the job list and scrap to get the attributes of the jobs*/
            foreach (IWebElement job in jobs)
            {
                string str_title, str_company, str_location, str_keywords, str_link;
                IWebElement elem_job_title = job.FindElement(By.CssSelector("#search-result-body > div > ul > li > span.job-info > a > h2"));
                str_title = elem_job_title.Text;

                IWebElement elem_company = job.FindElement(By.CssSelector("#search-result-body > div > ul > li > span.job-info > span.job-company"));
                str_company = elem_company.Text;

                IWebElement elem_location = job.FindElement(By.CssSelector("#search-result-body > div > ul > li > span.job-info > span.job-details > span.job-location > span"));
                str_location = elem_location.Text;

                IWebElement elem_keywords = job.FindElement(By.CssSelector("#search-result-body > div > ul > li > span.job-info > span.job-keywords"));
                str_keywords = elem_keywords.Text;

                IWebElement elem_job_link = job.FindElement(By.CssSelector("#search-result-body > div > ul > li > span.job-info > a"));
                str_link = elem_job_link.GetAttribute("href");

                Console.WriteLine("******* Job " + jcount + " *******");
                Console.WriteLine("Job title: " + str_title);
                Console.WriteLine("Company: " + str_company);
                Console.WriteLine("Location: " + str_location);
                Console.WriteLine("Keywords: " + str_keywords);
                Console.WriteLine("Link: " + str_link);
                Console.WriteLine("\n");

                // create newline to append to file
                string newLine = string.Format("{0}, {1}, {2}, {3}, {4}, {5}",
                    jcount.ToString(), str_title, str_company, str_location, str_keywords, str_link);
                output.AppendLine(string.Join(separator, newLine));

                // add new job to jsonfile
                _ictdata.Add(new ictdata()
                {
                    Id = jcount,
                    Title = str_title,
                    Company = str_company,
                    Location = str_location,
                    Keywords = str_keywords,
                    Link = str_link
                });

                jcount++;
                if (jcount > 5)
                {
                    string json = JsonSerializer.Serialize(_ictdata);
                    File.WriteAllText(@"C:\Users\finnd\source\repos\files\ictjobs.json", json);
                    break;
                }

            }
            // write all lines to file and give error if something goes wrong
            try
            {
                File.AppendAllText(ict_file, output.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine("Data could not be written to the CSV file.");
                return;
            }
            Console.WriteLine("Scraping Data from the 5 newest jobs on ictjobs.be completed and writen to file.");
        }

        // Swimrankings Scraping
        public class swimdata
        {
            public string Event { get; set; }
            public string Lane { get; set; }
            public string Time { get; set; }
            public string Date { get; set; }
            
        }

        public void SwimrankingsScraping(IWebDriver? driver)
        {
            driver.Url = url;

            // create a csv with headings
            String swim_csv = @"C:\Users\finnd\source\repos\files\swim_output.csv";
            String separator = ";";
            StringBuilder output = new StringBuilder();
            String[] headings = { "EventID", "Event", "LaneType", "Time", "Date" };
            output.AppendLine(string.Join(separator, headings));

            // create list to add to jsonfile
            List<swimdata> _swimdata = new List<swimdata>();

            /* Explicit Wait to ensure that the page is loaded completely by reading the DOM state */
            var timeout = 10000; /* Maximum wait time of 10 seconds */
            var wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(timeout));
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

            // implement search option
            Console.WriteLine("What is the last name of the swimmer?");
            string inputLastName = Console.ReadLine();

            // search lastname
            var searchOption = driver.FindElement(By.XPath("//*[@id=\"athlete_lastname\"]"));
            searchOption.SendKeys(inputLastName);
            searchOption.Submit();
            Console.WriteLine("Your search is processing...");

            Thread.Sleep(1000);

            Console.WriteLine("What is the first name of the swimmer?");
            string inputFirstName = Console.ReadLine();

            // search firstname
            var searchOption2 = driver.FindElement(By.XPath("//*[@id=\"athlete_firstname\"]"));
            searchOption2.SendKeys(inputFirstName);
            searchOption2.Submit();
            Console.WriteLine("Your search is processing...");

            Thread.Sleep(1000);

            // click correct swimmer 
            IWebElement swimmer_select = driver.FindElement(By.XPath("//*[@id=\"searchResult\"]/table/tbody/tr[2]/td[2]/a"));
            IJavaScriptExecutor mostrecentExecutor = (IJavaScriptExecutor)driver;
            mostrecentExecutor.ExecuteScript("arguments[0].click();", swimmer_select);

            Thread.Sleep(5000);

            By elem_swims = By.CssSelector("#content > table > tbody > tr > td > table.athleteBest > tbody > tr:nth-child(n+2)");
            ReadOnlyCollection<IWebElement> swims = driver.FindElements(elem_swims);

            /* Go through the event list and scrap to get the attributes of the events*/
            foreach (IWebElement swim in swims)
            {
                string str_event, str_lane, str_time, str_date;
                Console.WriteLine("The personal bests for swimmer " + inputFirstName + " " + inputLastName + " are:");

                IWebElement elem_event = swim.FindElement(By.CssSelector("#content > table > tbody > tr > td > table.athleteBest > tbody > tr > td.event > a"));
                str_event = elem_event.Text;
                Console.WriteLine("Event: " + str_event);

                IWebElement elem_lane = swim.FindElement(By.CssSelector("#content > table > tbody > tr > td > table.athleteBest > tbody > tr > td.course"));
                str_lane = elem_lane.Text;
                Console.WriteLine("lane-distance: " + str_lane);

                IWebElement elem_time = swim.FindElement(By.CssSelector("#content > table > tbody > tr > td > table.athleteBest > tbody > tr > td.time > a"));
                str_time = elem_time.Text;
                Console.WriteLine("time: " + str_time);

                IWebElement elem_date = swim.FindElement(By.CssSelector("#content > table > tbody > tr > td > table.athleteBest > tbody > tr > td.date"));
                str_date = elem_date.Text;
                Console.WriteLine("date: " + str_date);
                Console.WriteLine();
                scount++; 

                // create newline to append to file
                string newLine = string.Format("{0}, {1}, {2}, {3}, {4}",
                    scount.ToString(), str_event, str_lane, str_time, str_date);
                output.AppendLine(string.Join(separator, newLine));

                // add new event to jsonfile
                _swimdata.Add(new swimdata()
                {
                    Event = str_event,
                    Lane = str_lane,
                    Time = str_time,
                    Date = str_date,
                });
            }
            // write all events to json file
            string json = JsonSerializer.Serialize(_swimdata);
            File.WriteAllText(@"C:\Users\finnd\source\repos\files\swimdata.json", json);

            // write all lines to file and give error if something goes wrong
            try
            {
                File.AppendAllText(swim_csv, output.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine("Data could not be written to the CSV file.");
                return;
            }
            Console.WriteLine("Scraping Data from Swimrankings and writen to file done.");

        }

        public void Close_Browser(IWebDriver? driver)
        {
            driver.Quit();
        }
    }
}