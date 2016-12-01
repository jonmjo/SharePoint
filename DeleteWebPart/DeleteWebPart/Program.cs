using System;
using System.Linq;
using Microsoft.SharePoint;
using System.Collections.Generic;

namespace DeleteWebPart
{
    class Program
    {
        static void Main(string[] args)
        {
            bool listOnly = false;
            Console.ForegroundColor = ConsoleColor.White;

            if (args.Length == 2 && args[1].ToLower() == "/list") listOnly = true; ;

            if (args.Length != 2)
            {
                Console.WriteLine("Deletes a web part from the web part gallary.");
                Console.ResetColor();
                Console.WriteLine(
                    string.Format(
                        "Usage: {0} {1} {2}",
                        System.AppDomain.CurrentDomain.FriendlyName,
                        "http://localhost:51001",
                        "name.webpart"
                    )
                );
                Console.WriteLine(
                    string.Format(
                        "Usage: {0} {1} {2}",
                        System.AppDomain.CurrentDomain.FriendlyName,
                        "http://localhost:51001",
                        "webpartID"
                    )
                );

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("\nTo show all installed web parts.");
                Console.ResetColor();
                Console.WriteLine(string.Format("{0} {1}", System.AppDomain.CurrentDomain.FriendlyName, "/list" ) );
                
                return;
            }

            string siteUrl = args[0];
            string webpartName = listOnly ? string.Empty : args[1].ToLower();
            int webpartID = -1;
            int.TryParse(webpartName, out webpartID);

            using (SPSite parentSite = new SPSite(siteUrl))
            {
                List<int> toDelete = new List<int>();
                SPList list = parentSite.GetCatalog(SPListTemplateType.WebPartCatalog);

                Console.WriteLine("Searching...");
                Console.ResetColor();
                foreach (SPListItem item in list.Items)
                {
                    if (listOnly)
                    {
                        Console.WriteLine("Found: " + item.ID + ", " + item.Name);
                        continue;
                    }

                    if (webpartID > -1)
                    {
                        if (item.ID == webpartID)
                        {
                            Console.WriteLine("Found: " + item.Name + " with ID " + item.ID);
                            toDelete.Add(item.ID);
                        }
                    }
                    else
                    {
                        if (item["Web Part"].ToString().ToLower() == webpartName)
                        {
                            Console.WriteLine("Found: " + webpartName + " with ID " + item.ID);
                            toDelete.Add(item.ID);
                        }
                    }
                }

                if (listOnly) return;

                if (toDelete.Count() == 0) Console.WriteLine("Found nothing matching.");

                foreach(int i in toDelete)
                {
                    SPListItem item = list.GetItemById(i);
                    item.Delete();
                    Console.WriteLine("Deleted ID: " + i);
                }
                Console.WriteLine("Updating list...");
                list.Update();
                Console.WriteLine("Done.");
            }
        }
    }
}
