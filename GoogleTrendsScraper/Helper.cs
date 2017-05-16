using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Cache;
using System.Threading;
using System.Threading.Tasks;
using CefSharp;
using CefSharp.OffScreen;

namespace GoogleTrendsScraper
{
	/// <summary>
	/// A helper class that contains generic functions for all the scrapers
	/// </summary>
	internal class Helper
	{
		/// <summary>
		/// Path to the scraper's results folder
		/// </summary>
		public static string GoogleResultPath = "Results";

		/// <summary>
		/// Path to the scraper's screenshots folder
		/// </summary>
		private const string ScreenshotPath = "Screenshots";

		/// <summary>
		/// Path to the scraper's cache folder
		/// </summary>
		private const string CachePath = "Cache";

		/// <summary>
		/// Initialize the scraper variables and cleanup result directories
		/// </summary>
		public static void Initialize()
		{
			//set thread culture, for English error messages
			var culture = CultureInfo.CreateSpecificCulture("en-US");

			CultureInfo.DefaultThreadCurrentCulture = culture;
			CultureInfo.DefaultThreadCurrentUICulture = culture;

			Thread.CurrentThread.CurrentCulture = culture;
			Thread.CurrentThread.CurrentUICulture = culture;

			//cleanup output directories
			DeleteAllFiles(ScreenshotPath);
			DeleteAllFiles(GoogleResultPath);

			//initialize the browser engine
			var settings = new CefSettings { CachePath = CachePath };
			Cef.Initialize(settings, true, null);
		}

		/// <summary>
		/// Cut part of an image and save it as a separate image
		/// </summary>
		/// <param name="rect"> New image dimensions.</param>
		/// <param name="fileFrom">Source image path</param>
		/// <param name="fileTo">Destination image path</param>
		public static void CutImage(int[] rect, string fileFrom, string fileTo)
		{
			//generate an image
			var selection = new Rectangle(rect[0], rect[1], rect[2], rect[3]);
			var source = Image.FromFile(fileFrom);
			var bmp = new Bitmap(selection.Width, selection.Height);

			using (var gr = Graphics.FromImage(bmp))
			{
				gr.DrawImage(source, new Rectangle(0, 0, bmp.Width, bmp.Height), selection, GraphicsUnit.Pixel);
			}

			//and save it
			bmp.Save(fileTo, ImageFormat.Png);
		}

		/// <summary>
		/// Download a file and save it to a path
		/// </summary>
		/// <param name="url"> File URL.</param>
		/// <param name="path">File path</param>
		/// <returns>true if file was downloaded successfully, false otherwise</returns>
		public static bool DownloadFile(string url, string path)
		{
			try
			{
				if (url.IndexOf("http", StringComparison.Ordinal) != 0)
				{
					url = "http:" + url;
				}

				// Create WebRequest, and set headers
				var request = WebRequest.Create(url);
				var noCachePolicy = new HttpRequestCachePolicy(HttpRequestCacheLevel.NoCacheNoStore);
				request.CachePolicy = noCachePolicy;

				var buffer = new byte[1024];

				// Open stream and download the file
				using (var webResponse = (HttpWebResponse)request.GetResponse())
				{
					var fileStream = File.OpenWrite(path);
					using (var input = webResponse.GetResponseStream())
					{
						// ReSharper disable once PossibleNullReferenceException
						input.ReadTimeout = 5000;
						// ReSharper disable once PossibleNullReferenceException
						var size = input.Read(buffer, 0, buffer.Length);
						while (size > 0)
						{
							fileStream.Write(buffer, 0, size);
							size = input.Read(buffer, 0, buffer.Length);
						}
					}

					fileStream.Flush();
					fileStream.Close();

					return true;
				}
			}
			catch (Exception)
			{
				return false;
			}
		}

		/// <summary>
		/// Run a Javascript snippet on a browser view without returning any results
		/// </summary>
		/// <param name="view">Browser view for the snippet to be executed on</param>
		/// <param name="js">Javascript snippet</param>
		public static async Task PerformJs(ChromiumWebBrowser view, string js)
		{
			await view.EvaluateScriptAsync(js);
		}

		/// <summary>
		/// Run a Javascript snippet on a browser view, returning results as a string
		/// </summary>
		/// <param name="view">Browser view for the snippet to be executed on</param>
		/// <param name="js">Javascript snippet</param>
		/// <returns>Javascript script result, or "" if undefined</returns>
		public static async Task<string> GetJs(ChromiumWebBrowser view, string js)
		{
			var value = (await view.EvaluateScriptAsync(js)).Result;

			return value == null ? "" : value.ToString();
		}

		/// <summary>
		/// Get an element's coordinates in the browser window
		/// </summary>
		/// <param name="view">Browser view that has the element</param>
		/// <param name="js">Javascript item selector</param>
		/// <returns>A rectangle with the element's coordinates</returns>
		public static async Task<Rectangle> GetCoordinates(ChromiumWebBrowser view, string js)
		{
			// Get the bounding rectangle for an element
			var rect = (await GetJs(view, "var a = " + js + ".getBoundingClientRect();" +
							"a.left + '|' + a.top + '|' + a.width + '|' + a.height")).Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

			// Remove the fractional part
			for (var i = 0; i < rect.Length; i++)
			{
				if (rect[i].IndexOf(".", StringComparison.Ordinal) != -1)
				{
					rect[i] = rect[i].Substring(0, rect[i].IndexOf('.', 0));
				}
			}
			
			// And return the resulting rectangle
			var result = new Rectangle(Convert.ToInt32(rect[0]), Convert.ToInt32(rect[1]), Convert.ToInt32(rect[2]), Convert.ToInt32(rect[3]));
			return result;
		}

		/// <summary>
		/// Take a screenshot of a browser view and save it to a folder
		/// </summary>
		/// <param name="view">Browser view to take a screenshot</param>
		/// <returns>Screenshot path</returns>
		public static async Task<string> TakeScreenshot(ChromiumWebBrowser view)
		{
			var screenshotGuid = Guid.NewGuid();

			var screenshotPath = ScreenshotPath + "/" + screenshotGuid + ".png";

			using (var task = await view.ScreenshotAsync())
			{
				task.Save(screenshotPath);
			}

			return screenshotPath;
		}

		/// <summary>
		/// Deletes all files in a target folder
		/// </summary>
		/// <param name="path">Path to the folder</param>
		public static void DeleteAllFiles(string path)
		{
			var di = new DirectoryInfo(path);
			foreach (var file in di.GetFiles())
			{
				file.Delete();
			}
		}

		/// <summary>
		/// Write a string to the console, with current timestamp
		/// </summary>
		public static void Log(string text)
		{
			var output = DateTime.Now + " " + text;
			Console.WriteLine(output);
		}
	}
}
