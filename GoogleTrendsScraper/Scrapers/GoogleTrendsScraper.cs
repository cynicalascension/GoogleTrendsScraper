using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using CefSharp;
using CefSharp.OffScreen;
using OfficeOpenXml;

namespace GoogleTrendsScraper.Scrapers
{
	internal class GoogleTrendsScraper
	{
		/// <summary>
		/// Main Google Trends URL
		/// </summary>
		private const string GoogleTrendsUrl = "https://trends.google.com/trends/";

		/// <summary>
		/// A flag that indicates whether scraping is over or not
		/// </summary>
		private static bool SiteReady;

		/// <summary>
		/// A list with all results scraped so far
		/// </summary>
		private static List<GoogleResult> GoogleResultList;

		/// <summary>
		/// A single scraped result from GoogleTrends
		/// </summary>
		private class GoogleResult
		{
			/// <summary>
			/// Story title
			/// </summary>
			public string Title;

			/// <summary>
			/// Path to the story image, if available
			/// </summary>
			public string ImgPath;

			/// <summary>
			/// Path to the story google popularity chart, if available
			/// </summary>
			public string ImgPathChart;

			/// <summary>
			/// Story URL on Google
			/// </summary>
			public string GoogleUrl;

			/// <summary>
			/// Story URL on an external site
			/// </summary>
			public string ExternalUrl;

			/// <summary>
			/// Story's main image URL
			/// </summary>
			public string ImgUrl;
		}

		/// <summary>
		/// Scrape the Google Trends title page
		/// </summary>
		public static async Task ScrapeGoogle()
		{
			GoogleResultList = new List<GoogleResult>();

			using (var browser = new ChromiumWebBrowser(GoogleTrendsUrl))
			{
				browser.Size = new Size(2000, 2000);

				SiteReady = false;
			
				// Create the load complete handler 
				EventHandler<LoadingStateChangedEventArgs> handler = null;

				handler = async (sender, args) =>
				{
					if (args.IsLoading) return;
					var view = (ChromiumWebBrowser)sender;
					view.LoadingStateChanged -= handler;

					// Wait 10 seconds, just in case, for all AJAX calls to complete
					await Task.Delay(10000);

					Helper.Log("Google Trends page loaded successfully");

					var pageRenderPath = await Helper.TakeScreenshot(view);

					var windowHeight = await Helper.GetJs(view, "window.innerHeight;");

					// For the first 50 trends
					for (var i = 0; i < 50; i++)
					{
						// Get each trend's properties
						var googleStoryLink = await Helper.GetJs(view, "document.getElementsByClassName('trending-story-wrapper')[" + i + "].getElementsByClassName('trending-story ng-isolate-scope')[0].getAttribute('ng-href');");

						var storyTitle = await Helper.GetJs(view, "document.getElementsByClassName('trending-story-wrapper')[" + i + "].getElementsByClassName('ng-binding')[0].innerHTML;");

						var externalStoryLink = await Helper.GetJs(view, "document.getElementsByClassName('trending-story-wrapper')[" + i + "].getElementsByClassName('image-wrapper ng-scope')[0].getAttribute('ng-href');");

						var imgLink = await Helper.GetJs(view, "document.getElementsByClassName('trending-story-wrapper')[" + i + "].getElementsByClassName('image fe-atoms-generic-hide-in-mobile ng-scope')[0].getAttribute('src');");

						var newImgPath = "none";

						if (imgLink != "undefined")
						{
							// Download the story image, if it is available
							newImgPath = Helper.GoogleResultPath + "/" + i + "_image.png";
							Helper.DownloadFile(imgLink, newImgPath);
						}

						var lowerline = Convert.ToInt32(windowHeight);

						// Get the story image coordinates
						var storyImgCoord = await Helper.GetCoordinates(view, "document.getElementsByClassName('trending-story-wrapper')[" + i + "].getElementsByClassName('sparkline-chart ng-scope')[0]");

						// If the story image is not visible
						if (storyImgCoord.Top + storyImgCoord.Height + 20 > lowerline)
						{
							// Scroll down to it
							await Helper.PerformJs(view, "window.scrollTo(0, window.scrollY + 2000);");

							await Task.Delay(1000);

							while (view.IsLoading)
							{
								await Task.Delay(10);
							}

							// Wait 10 seconds, just in case, for all AJAX calls to complete
							await Task.Delay(10000);

							Helper.Log("Page scrolled down successfully.");

							// And re-take the page screenshot and image coordinates
							storyImgCoord = await Helper.GetCoordinates(view, "document.getElementsByClassName('trending-story-wrapper')[" + i + "].getElementsByClassName('sparkline-chart ng-scope')[0]");

							pageRenderPath = await Helper.TakeScreenshot(view);
						}

						// Cut out the chart image from the page screenshot
						var chartImgPath = Helper.GoogleResultPath + "/" + i + "_chart.png";
						Helper.CutImage(new[] { storyImgCoord.Left, storyImgCoord.Top, storyImgCoord.Width, storyImgCoord.Height }, pageRenderPath, chartImgPath);

						var result = new GoogleResult
						{
							GoogleUrl = googleStoryLink,
							Title = storyTitle,
							ExternalUrl = externalStoryLink,
							ImgUrl = imgLink,
							ImgPathChart = chartImgPath,
							ImgPath = newImgPath
						};

						GoogleResultList.Add(result);
					}

					SiteReady = true;
				};

				// And link the handler to the loading state changed event
				browser.LoadingStateChanged += handler;

				// Wait for the scraping to complete
				while (SiteReady != true)
				{
					await Task.Delay(10);
				}
			}
		}

		/// <summary>
		/// Save all scraped results to an XLS file
		/// </summary>
		public static void SaveGoogle()
		{
			var reportFile = Helper.GoogleResultPath + "/GoogleTrendsOutput.xlsx";

			// Create the file
			var newFile = new FileInfo(reportFile);

			using (var pck = new ExcelPackage(newFile))
			{
				// And add a sheet to it
				var ws = pck.Workbook.Worksheets.Add("Content");

				// Write the headers
				ws.Cells[1, 1].Value = "Title";
				ws.Cells[1, 2].Value = "ImgPath";
				ws.Cells[1, 3].Value = "ImgPathChart";
				ws.Cells[1, 4].Value = "GoogleUrl";
				ws.Cells[1, 5].Value = "ExternalUrl";
				ws.Cells[1, 6].Value = "ImgUrl";

				// Then write the body itself in a loop
				for (var p = 0; p < GoogleResultList.Count; p++)
				{
					ws.Cells[p + 2, 1].Value = GoogleResultList[p].Title;
					ws.Cells[p + 2, 2].Value = GoogleResultList[p].ImgPath;
					ws.Cells[p + 2, 3].Value = GoogleResultList[p].ImgPathChart;
					ws.Cells[p + 2, 4].Value = GoogleResultList[p].GoogleUrl;
					ws.Cells[p + 2, 5].Value = GoogleResultList[p].ExternalUrl;
					ws.Cells[p + 2, 6].Value = GoogleResultList[p].ImgUrl;
				}

				// Auto-size the columns to make the result look prettier
				ws.Cells[ws.Dimension.Address].AutoFitColumns();

				// And save the results
				pck.Save();
			}
		}
	}
}
