using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Publishing;

namespace AvailablePageLayoutInSite
{
    class Program
    {
        static void Main(string[] args)
        {
            string url = "http://localhost:51001";
            string pageLayoutToRemove = "ChalmersJobOpeningsPage.aspx";
            try
            {
                Console.Write("\n Will search for PageLayout " + pageLayoutToRemove);

                using (SPSite oSite = new SPSite(url))
                {
                    using (SPWeb oWeb = oSite.OpenWeb("/"))
                    {
                        List<SPWeb> webs = new List<SPWeb>();
                        var swc = removePageLayoutFromWeb(oWeb, pageLayoutToRemove);
                        foreach (SPWeb s in swc) webs.Add(s);


                        for (int i = 0; i < webs.Count; i++)
                        {
                            Console.Write("\n Doing " + (i+1) + " of " + webs.Count);
                            var moreSWC = removePageLayoutFromWeb(webs[i], pageLayoutToRemove);

                            foreach (SPWeb addWb in moreSWC) webs.Add(addWb);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Console.Write("An Error occured " + ex.Message);
            }
            Console.Write("\nPress any key to end the program...");
            Console.ReadKey(true);
        }

        public static SPWebCollection removePageLayoutFromWeb(SPWeb oWeb, string pageLayoutToRemove)
        {
            PublishingWeb pWeb = PublishingWeb.GetPublishingWeb(oWeb);

            Console.Write("\n Doing " + oWeb.Url);

            if (!pWeb.IsInheritingAvailablePageLayouts)
	        {
		        Console.Write("\n  Does not inherit. Searching...");
		        
		        List<PageLayout> myArray = new List<PageLayout>();
		        
		
		        var availablePageLayouts = pWeb.GetAvailablePageLayouts();
		        Console.Write("\n " + availablePageLayouts.Length.ToString() + " Page Layouts assoiciated with web.");
		        if (availablePageLayouts.Length > 0)
		        {
			        for(int i = 1; i < availablePageLayouts.Length; i++)
			        {
                        Console.Out.Flush();
				        string strapl = availablePageLayouts[i].Name;
				
				        if (strapl.Length == 0)
				        {
					        Console.Write("\nPage Layout without filename. Skipping " + availablePageLayouts[i].Title);
				        }
				        else
				        {
					        Console.Write("\n   Comparing " + strapl + " with " + pageLayoutToRemove);
                            
					        if (strapl.CompareTo(pageLayoutToRemove) == 0)
					        {
						        var strrh = availablePageLayouts[i].Title;
                                Console.Write("\n   Removing page" + strapl + ". Press Enter to continue.");
						        Console.ReadLine();
						        
					        }
					        else
					        {
						        Console.Write("\n   Keeping page" + strapl);
						        myArray.Add(availablePageLayouts[i]);
					        }
				        }
			        }
		
			        if (myArray.Count > 0)
			        {
				        Console.Write("\n  Updating Web");
				        
				        pWeb.SetAvailablePageLayouts(myArray.ToArray(), false);
				        pWeb.Update();
			        }
			        else
			        {
				        Console.Write("\n  No pagelayouts to set. Skipping update. (At least one PageLayout must be left to do update).");
			        }
		        }
		
	        }
	        else
	        {
		        Console.Write("\n  Inherits. Skipping.");
	        }
	        Console.Write("\n");
            return oWeb.Webs;
        }

    }
}


