using System;
using System.Linq;
using Microsoft.SharePoint;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace GetSharepointLists
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: https://localhost:51001 MyList.csv /r");
                Console.WriteLine("/r is case sensitive and optional.");
                return;
            }

            bool recursive = args.Contains("/r");

            Uri uri = new Uri(args.First(item => item.StartsWith("http")));
            var targetFile = args.First(item => item.StartsWith("http") == false && item.Equals("/r") == false);
            GetLists(recursive, uri, targetFile);
        }

        private static void GetLists(bool recursive, Uri uri, string targetFile)
        {
            List<ListDataRow> csvFileItems = new List<ListDataRow>();

            var itm = new ListDataRow();
            csvFileItems.Add(itm);

            itm.ItemType = "Item type";
            itm.BaseType = "Type";
            itm.Title = "Title";
            itm.ListTypeID = "List type ID";
            itm.Created = "Created";
            itm.Editor = "Editor";
            itm.LastEntry = "Last entry";
            itm.itemCount = "Items in list";
            itm.Url = "URL";


            var webs = new List<SPWeb>();

            using (SPSite oSite = new SPSite(uri.OriginalString))
            {
                using (SPWeb oWeb = oSite.OpenWeb(uri.PathAndQuery))
                {
                    webs.Add(oWeb);

                    for (int i = 0; i < webs.Count; i++)
                    {
                        Console.WriteLine("Processing " + (i+1) + "/" + webs.Count + ": " + oWeb.Site.Url + webs[i].ServerRelativeUrl);
                        var foundWebs = AddLists(webs[i], csvFileItems);
                        if (recursive) webs.AddRange(foundWebs);
                    }

                }
            }

            var rows = csvFileItems
                .Select(x => x.ItemType + ";" + x.Created + ";" + x.Title + ";" + x.LastEntry + ";" + x.BaseType + ";" + x.ListTypeID + ";" + x.Editor + ";" +x.itemCount + ";" + x.Url)
                .ToArray();
            File.WriteAllText(targetFile, string.Join(Environment.NewLine, rows));
            Console.WriteLine("Saved file to: " + targetFile);
        }

        private static List<SPWeb> AddLists(SPWeb oWeb, List<ListDataRow> csvFileItems)
        {
            foreach (SPList list in oWeb.Lists)
            {
                var itm = new ListDataRow();
                csvFileItems.Add(itm);

                itm.ItemType = "List";
                itm.BaseType = list.BaseType.ToString();
                itm.Title = list.Title;
                itm.Created = list.Created.ToShortDateString();
                itm.Editor = list.Author.LoginName;
                itm.LastEntry = list.LastItemModifiedDate.ToShortDateString();
                itm.ListTypeID = list.TemplateFeatureId.ToString();
                itm.itemCount = list.ItemCount.ToString();

                itm.Url = oWeb.Site.Url + list.DefaultViewUrl;

                AddFormList(list, csvFileItems);
            }
            return oWeb.Webs.ToList();
        }

        private static void AddFormList(SPList list, List<ListDataRow> csvFileItems)
        {
            foreach (SPForm form in list.Forms)
            {
                var itm = new ListDataRow();
                csvFileItems.Add(itm);

                itm.ItemType = "Form";
                itm.BaseType = form.Type.ToString();
                itm.Title = form.ParentList.Title;
                itm.Created = form.ParentList.Created.ToShortDateString();
                itm.Editor = form.ParentList.Author.LoginName;
                itm.LastEntry = form.ParentList.LastItemModifiedDate.ToShortDateString();
                itm.ListTypeID = form.TemplateName;
                itm.itemCount = form.ParentList.ItemCount.ToString();

                itm.Url = form.ParentList.ParentWeb.Site.Url + form.ServerRelativeUrl;
            }
        }
    }
}
