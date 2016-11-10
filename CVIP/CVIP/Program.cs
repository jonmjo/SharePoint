using System;
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Publishing;

namespace CVIP
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 4)
            {
                Console.WriteLine(
                    string.Format(
                        "Exempel:\n{0} {1} {2} {3} {4}",
                        System.AppDomain.CurrentDomain.FriendlyName,
                        "http://localhost:51001",
                        "/sv/projekt",
                        "ShowProjectTab",
                        "true"
                    )
                );
                return;
            }

            string siteUrl = args[0];
            string dir = args[1];
            string fieldName = args[2]; //"ShowProjectTab";
            string value = args[3];
            int type=0;

            bool parsedBool;
            int parsedInt = 0;

            if (bool.TryParse(value, out parsedBool)) type = 1;
            else if (int.TryParse(value, out parsedInt)) type = 2;
            else throw new Exception("Only designed to take arguments in the form of integers and booleans.");
            
            Console.WriteLine("Starting...");
            
            using (SPSite spsite = new SPSite(siteUrl)) // http://localhost:51001
            {
                using (SPWeb staffWeb = spsite.OpenWeb(dir)) // site.OpenWeb(stafWebId))
                {
                    PublishingWeb pweb = PublishingWeb.GetPublishingWeb(staffWeb);
                    SPList staffSitePagesList = pweb.PagesList;

                    SPListItemCollection col = staffSitePagesList.Items;

                    //SPField field = (staffSitePagesList).Fields.GetField(fieldName);

                    Console.WriteLine("Found " + col.Count + " items.");
                    int n = 1;
                    
                    foreach (SPListItem item in col)
                    {
                        try
                        {
                            Console.WriteLine(n++ + " / " + col.Count + ": " + item.Url);
                            
                            //if ((bool)item[fieldName]) continue;
                            if (type == 1)
                            {
                                if ((bool)item[fieldName] != parsedBool)
                                {
                                    item[fieldName] = parsedBool;
                                    item.SystemUpdate(false);
                                }
                            }
                            else if (type == 2)
                            {
                                if ((int)item[fieldName] != parsedInt)
                                {
                                    item[fieldName] = parsedInt;
                                    item.SystemUpdate(false);
                                }

                            }
                            else throw new Exception("Only designed to take arguments in the form of integers and booleans.");
                            
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine(exception.Message);
                        }
                    }
                }
            }
            Console.WriteLine("Program ended normally.");
        }
    }
}
