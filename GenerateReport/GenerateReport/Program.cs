using System;
using System.Linq;
using Microsoft.SharePoint;
using System.Text.RegularExpressions;
using Chalmers.Core.Repositories;
using System.IO;
using Chalmers.Core;
using System.Collections.Generic;

namespace GenerateReport
{
    class Program
    {
        private static readonly char[] urlTrimChars = { '/' };
        private static readonly Regex webRelativeUrlRegex = new Regex("/[^/]+/[^/]+.aspx$", RegexOptions.Compiled & RegexOptions.IgnoreCase);

        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.White;

            if (args.Length < 2 || args.Length > 3)
            {
                Console.WriteLine("Generates reports.");
                Console.ResetColor();
                Console.WriteLine(
                    string.Format(
                        "Usage: {0} {1} {2}",
                        System.AppDomain.CurrentDomain.FriendlyName,
                        "http://localhost:51001",
                        "/sv/utbildning"
                    )
                );
                Console.WriteLine(
                    string.Format(
                        "Usage: {0} {1} {2} {3}",
                        System.AppDomain.CurrentDomain.FriendlyName,
                        "http://localhost:51001",
                        "/sv/utbildning",
                        "Page_responsible_email"
                    )
                );

                return;
            }

            string host = args[0];
            string path = args[1];
            DateTime start = DateTime.Now;

            if (args.Length == 2)
            {
                GenerateFilteredPageReport(host, path, null, null);
            }
            else if (args.Length == 3)
            {
                GenerateUserPageReport(host, path, args[2]);
            }
            else
            {
                throw new Exception("Not impmlemented.");
            }
            DateTime finied = DateTime.Now;
            Console.WriteLine(
                string.Format("Started at:  {0}.\nFinished at: {1}.\nTotal:       {2}", 
                    start, 
                    finied, 
                    finied.Subtract(start)));
        
#if DEBUG
            Console.Write("Press any key to exit program...");
            Console.ReadKey();
#endif
        }


        public static void GenerateUserPageReport(string host, string path, string email)
        {
            email = email ?? string.Empty;
            email = email.Trim();
            GenerateFilteredPageReport(host, path, "UserPageReport", (p, c) => email.Equals(c, StringComparison.InvariantCultureIgnoreCase));
        }

        public static void GenerateFilteredPageReport(string host, string path, string reportName, Func<PageEntity, string, bool> rowFilter)
        {
            if (!(path.StartsWith("/sv") || path.StartsWith("/en")))
            {
                Console.WriteLine("URL '" + path + "' not valid. Must start with /sv or /en");
                return;
            }

            using (SPSite oSite = new SPSite(host))
            {
                using (SPWeb oWeb = oSite.OpenWeb(path))
                {
                    reportName = reportName ?? "PageReport";
                    rowFilter = rowFilter ?? ((p, c) => true);

                    string langCode = path.Substring(1, 2);
                    string langRootUrl = string.Format("/{0}", langCode);
                    Uri siteUri = new Uri(oSite.Url);
                    DateTime datestamp = DateTime.Now;
                    string fileName = string.Format("{0}-{1}-{2}-{3:yyyy-MM-dd-HH-mm-ss}.txt", reportName, siteUri.Host, langCode, datestamp);

                    ContactPersonRepository contactRepository = new ContactPersonRepository(oWeb) { ServerRelativeWebUrl = langRootUrl };

                    using (CsvFileWriter writer = new CsvFileWriter(fileName, ';'))
                    {
                        CsvRow headerRow = new CsvRow();
                        headerRow.Add("Title");
                        headerRow.Add("URL");
                        headerRow.Add("Page responsible");
                        headerRow.Add("Created");
                        headerRow.Add("Modified");
                        writer.WriteRow(headerRow);

                        using (SPWeb langRootWeb = oSite.OpenWeb(path))
                        {
                            List<SPWeb> websToProcess = new List<SPWeb>();
                            websToProcess.Add(langRootWeb);

                            for (int i = 0; i < websToProcess.Count(); i++)
                            {
                                Console.WriteLine(string.Format("Processing web {0}/{1}: {2}", i+1, websToProcess.Count(), websToProcess[i].Name));
                                var moreWebs = GetPageReport(rowFilter, siteUri, contactRepository, writer, websToProcess[i]);
                                foreach (SPWeb w in moreWebs) websToProcess.Add(w);
                            }
                        }

                        writer.Flush();
                    }
                }
            }
        }

        private static SPWebCollection GetPageReport(Func<PageEntity, string, bool> rowFilter, Uri siteUri, ContactPersonRepository contactRepository, CsvFileWriter writer, SPWeb web)
        {
            SingleWebPagesRepository pagesRepository = new SingleWebPagesRepository() { Web = web };
            IEnumerable<PageEntity> pages = pagesRepository.GetItems(web.Site).Cast<PageEntity>();

            int sida = 0;
            foreach (PageEntity page in pages.OrderBy(x => x.Title))
            {
                ClearCurrentConsoleLine();
                Console.Write(string.Format(" Processing page {0}/{1}: {2}", ++sida, pages.Count(), page.FileRef));
                string pageRelativeUrl = page.FileRef;
                int indexSemicHash = pageRelativeUrl.IndexOf(";#");

                if (indexSemicHash >= 0 && pageRelativeUrl.Length > 2) 
                    pageRelativeUrl = pageRelativeUrl.Substring(indexSemicHash + 2);

                pageRelativeUrl = string.Format("/{0}", pageRelativeUrl.TrimStart(urlTrimChars));

                string webRelativeUrl = webRelativeUrlRegex.Replace(pageRelativeUrl, string.Empty);
                string contact = contactRepository.GetContactAddressForPage(pageRelativeUrl, webRelativeUrl, web.Site);

                if (!rowFilter(page, contact)) continue;

                Uri pageUri = new Uri(siteUri, pageRelativeUrl);
                CsvRow row = new CsvRow();
                row.Add(page.Title);
                row.Add(pageUri.ToString());
                row.Add(contact);
                row.Add(page.Created);
                row.Add(page.Modified);
                writer.WriteRow(row);
            }
            Console.WriteLine("\n");
            return web.Webs;
        }

        public static void ClearCurrentConsoleLine()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
        }

    }
}
