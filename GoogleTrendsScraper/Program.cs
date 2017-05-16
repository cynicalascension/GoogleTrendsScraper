using System.Threading.Tasks;

namespace GoogleTrendsScraper
{
	internal class Program
	{
		private static void Main()
		{
			//initialize the browser engine
			Helper.Initialize();

			Helper.Log("Starting scrape...");

			//scrape the Google Trends page
			Task.Run(async () => await Scrapers.GoogleTrendsScraper.ScrapeGoogle()).Wait();

			//save results
			Scrapers.GoogleTrendsScraper.SaveGoogle();

			Helper.Log("All done!");
		}
	}
}
