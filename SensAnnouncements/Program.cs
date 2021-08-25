using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using PuppeteerSharp;

namespace SensAnnouncements
{
    class Program
    {
        private const string SensAnnouncementsUrl = "https://clientportal.jse.co.za/communication/sens-announcements";
        private static Browser _browser;
        private static string _reportFileName = "SensReport.xlsx";

        static async Task Main(string[] args)
        {
            if (args.Any())
            {
                _reportFileName = args[0];
            }

            await new BrowserFetcher().DownloadAsync(BrowserFetcher.DefaultChromiumRevision);
            _browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using (_browser)
            {
                var rawAnnouncementListHtml = await ScrapeRawSensAnnouncements();
                var reportingData = await PrepareReportingData(rawAnnouncementListHtml);
                ReportHelper.FormatOutput(reportingData, _reportFileName);
            }
        }

        private static async Task<string> ScrapeRawSensAnnouncements()
        {
            var page = await _browser.NewPageAsync();
            await page.GoToAsync(SensAnnouncementsUrl);

            //Wait until SENS announcements are all loaded
            await page.WaitForFunctionAsync("() => !document.getElementById('loader')");
            var rawAnnouncementListHtml = await page.EvaluateFunctionAsync<string>("() => document.getElementById('announcements').innerHTML");
            return rawAnnouncementListHtml;
        }

        private static async Task<IReadOnlyCollection<SensAnnouncementReportingItem>> PrepareReportingData(string rawAnnouncements)
        {
            var page = await _browser.NewPageAsync();
            await page.SetContentAsync(rawAnnouncements);
            var sensAnnouncementsToReport = new List<SensAnnouncementReportingItem>();
            var sensAnnouncements = await page.QuerySelectorAllAsync("li");
            foreach (var announcement in sensAnnouncements)
            {
                var elementA = await announcement.QuerySelectorAsync("a");
                var title = await elementA.EvaluateFunctionAsync<string>("(el) => el.innerText");
                var linkToAnnouncement = await elementA.EvaluateFunctionAsync<string>("(el) => el.getAttribute('href')");

                var elementP = await announcement.QuerySelectorAsync("p");
                var announcementSidTimestamp = await elementP.EvaluateFunctionAsync<string>("(el) => el.innerText");

                sensAnnouncementsToReport.Add(new SensAnnouncementReportingItem(title, linkToAnnouncement, announcementSidTimestamp));
            }

            return sensAnnouncementsToReport;
        }
    }
}
