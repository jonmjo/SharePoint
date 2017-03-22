using System;
using System.Linq;
using Microsoft.SharePoint;
using System.Collections.Generic;
using System.Text;

namespace DoesPageExist
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.White;

            if (args.Length != 1)
            {
                Console.WriteLine("List checked out files in web.");
                Console.ResetColor();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("\nTo performe a check:");
                Console.ResetColor();
                Console.WriteLine(string.Format("{0} {1}",
                    System.AppDomain.CurrentDomain.FriendlyName,
                    "http://localhost:51001/sv/"));

                return;
            }

            Uri fullSiteUrl = new Uri(args[0]);
            string siteUrl = fullSiteUrl.AbsoluteUri.Replace(fullSiteUrl.AbsolutePath, string.Empty);

            using (SPSite oSite = new SPSite(siteUrl))
            using (SPWeb oWeb = oSite.OpenWeb(fullSiteUrl.AbsolutePath))
            {
                var pubList = Microsoft.SharePoint.Publishing.PublishingWeb.GetPublishingWeb(oWeb);
                SPDocumentLibrary library = (SPDocumentLibrary)pubList.PagesList;

                Console.WriteLine("Searching...");
                Console.ResetColor();

                // ...print information about files uploaded but not checked in.
                IList<SPCheckedOutFile> files = library.CheckedOutFiles;
                foreach (SPCheckedOutFile file in files)
                {
                    Console.WriteLine("Checked out to: {0}.", file.CheckedOutBy);
                    Console.WriteLine(" /{0}/{1}" + Environment.NewLine, file.DirName, file.LeafName);

                    // This is the code to check in the document
                    //file.TakeOverCheckOut();
                    //SPListItem docItem = library.GetItemById(file.ListItemId);
                    //docItem.File.CheckIn(string.Empty);
                    //docItem.File.Update();
                }

                Console.WriteLine("Done.");
            }
        }
    }
}
