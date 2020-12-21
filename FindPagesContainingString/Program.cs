using System;
using System.Linq;
using Microsoft.SharePoint;
using System.Collections.Generic;
using CamlexNET;
using System.IO;

namespace FindPagesContaingingString
{
    class Program
    {

        private static string lastValue = "";


        static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                Console.WriteLine("Parameter 1=URL. Exmpelvis http://localhost:51001/");
                Console.WriteLine("Parameter 2=Sträng att hitta. Exmpelvis chalmersinnovation.com");
                Console.WriteLine("Parameter 3=Skriv utdata till filnamn");
                Console.WriteLine("FindPagesContainingString.exe http://localhost:51001/ chalmersinnovation.com c:\\temp\\utfil.txt");
                return;
            }

            IList<SPWeb> webs = new List<SPWeb>();

            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    using (SPSite aSite = new SPSite(args[0]))
                    {
                        SPWeb aWeb = aSite.OpenWeb();

                        foreach (SPWeb web in aWeb.Webs) webs.Add(web);

                        for (int i = 0; i < webs.Count(); i++)
                        {
                            Console.WriteLine("Bearbetar: " + (i + 1) + " av " + webs.Count() + ". " + webs[i].Url);
                            SPWebCollection websFromWeb = processWeb(args, webs[i]);
                            foreach (SPWeb web in websFromWeb) webs.Add(web);
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine("Klar. Tryck på valfri tangent för att avsluta...");
            Console.ReadKey();
            
        }

        private static SPWebCollection processWeb(string[] args, SPWeb theWeb)
        {
            foreach (SPList list in theWeb.Lists)
            {

                foreach (SPListItem li in list.Items)
                {
                    for (int j = 0; j < li.Fields.Count - 1; j++)
                    {
                        try
                        {
                            if (li[j] != null && li[j].ToString().ToLower().Contains(args[1]))
                            {
                                string value = theWeb.Url + "/" + li.Url.ToString();
                                if (lastValue.Equals(value) == false)
                                {
                                    Console.WriteLine("- " + value);
                                    File.AppendAllText(args[2], value + Environment.NewLine);
                                    lastValue = value;
                                }
                            }
                        }
                        catch (IndexOutOfRangeException ex)
                        {
                        }
                        catch (ArgumentOutOfRangeException ex)
                        {
                        }
                    }
                }
            }
            return theWeb.Webs;
        }
    }
}
