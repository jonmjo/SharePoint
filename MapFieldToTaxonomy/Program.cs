using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;

namespace MapFieldToTaxonomy
{
    class Program
    {
		// This does not work
        static void Main(string[] args)
        {
            if (args.Length < 5 || args.Length > 6)
            {
                Console.WriteLine("Maps a SPField to a specific point in the Managed Metadata tree (Taxonomy tree).");
                Console.ResetColor();
                Console.WriteLine(
                    string.Format(
                        "Usage: {0} {1} {2} {3} {4} {5} {6}",
                        System.AppDomain.CurrentDomain.FriendlyName,
                        "http://localhost:51001",
                        "TaxChalmersDepartment",
                        "TaxChalmersDepartmentHiddenFieldID",
                        "Chalmers.se",
                        "\"Chalmers Institutions Enterprise Taxonomy\"",
                        "Departments"
                        
                    )
                );
                Console.WriteLine("The Term (Departments in the example above) is not mandatory.\nIf obmitted will the field be mapped to the termset.");
                Console.WriteLine(
                    string.Format(
                        "Usage: {0} {1} {2} {3} {4} {5} {6}",
                        System.AppDomain.CurrentDomain.FriendlyName,
                        "http://localhost:51001",
                        "TaxChalmersDepartment",
                        "{FE44AD5C-5DF8-4846-A623-A3C693FF07A5}",
                        "Chalmers.se",
                        "\"Chalmers Institutions Enterprise Taxonomy\"",
                        "Departments"
                    )
                );
                Console.WriteLine(
                    string.Format(
                        "Usage: {0} {1} {2} {3} {4} {5} {6}",
                        System.AppDomain.CurrentDomain.FriendlyName,
                        "http://localhost:51001",
                        "TaxChalmersDepartment",
                        "{FE44AD5C-5DF8-4846-A623-A3C693FF07A5}",
                        "Chalmers.se",
                        "\"Chalmers Institutions Enterprise Taxonomy\"",
                        ""
                    )
                );


                return;
            }

            Dictionary<string, string> suppliedParams = new Dictionary<string,string>();
            suppliedParams.Add("server", args[0]);
            suppliedParams.Add("fieldName", args[1]);
            suppliedParams.Add("hiddenFieldGUID", args[2]);

            suppliedParams.Add("group", args[3]);
            suppliedParams.Add("termSet", args[4]);
            if (args.Length == 6) suppliedParams.Add("term", args[5]);
            else suppliedParams.Add("term", args[4]);

            using (SPSite oSite = new Microsoft.SharePoint.SPSite(suppliedParams["server"]))
            {
                var session = new TaxonomySession(oSite);
                TermStore termStore = session.TermStores[0];
                Guid TaxChalmersDepartmentHiddenFieldID = new Guid(suppliedParams["hiddenFieldGUID"]);

                SetupTaxonomyField(
                    oSite, 
                    termStore,
                    suppliedParams["group"],
                    suppliedParams["termSet"],

                    suppliedParams["term"],
                    suppliedParams["fieldName"],
                    TaxChalmersDepartmentHiddenFieldID);
            }
        }


        static void SetupTaxonomyField(
            SPSite _site, 
            TermStore termStore,
            string groupName,
            string termSetName,

            string subTermName,
            string taxonomyFieldName,
            Guid taxonomyFieldNoteFieldId)
        {
            if (termStore == null)
            {
                return;
            }
            //TermStore termStore = GetTermStore(termStoreName);
            Group group = GetGroup(termStore, groupName);
            TermSet termSet = GetTermSet(group, termSetName);
            Term term = null;

            try
            {
                TaxonomyField taxonomyField = GetTaxonomyField(taxonomyFieldName, _site);
                taxonomyField.SspId = termStore.Id;
                taxonomyField.TermSetId = termSet.Id;

                if (subTermName.Equals(termSetName))
                {
                    taxonomyField.AnchorId = new Guid("{00000000-0000-0000-0000-000000000000}");
                }
                else
                {
                    term = GetTerm(termSet, subTermName);
                    taxonomyField.AnchorId = term.Id;
                }

                taxonomyField.TextField = taxonomyFieldNoteFieldId;
                taxonomyField.Update(true);
            }
            catch (Exception exception)
            {
                Console.BackgroundColor = System.ConsoleColor.Red;
                Console.WriteLine(exception.Message + exception.StackTrace);
                Console.ResetColor();
            }
        }

        static TaxonomyField GetTaxonomyField(string internalFieldName, SPSite _site)
        {
            return (TaxonomyField)_site.RootWeb.Fields.GetFieldByInternalName(internalFieldName);
        }

        static Term GetTerm(TermSet termSet, string termName)
        {
            if (termSet != null)
            {
                foreach (Term term in termSet.Terms)
                {
                    if (term.GetDefaultLabel(1033) == termName)
                    {
                        return term;
                    }
                }
            }
            return null;

        }
        static Group GetGroup(TermStore termStore, string groupName)
        {
            var group = termStore.Groups.FirstOrDefault(g => g.Name.Equals(groupName, StringComparison.OrdinalIgnoreCase));
            if (group == null)
            {
                Console.BackgroundColor = System.ConsoleColor.Red;
                Console.WriteLine(
                    String.Format("The term store group {0} could not be found.", groupName)
                );
                Console.ResetColor();
            }

            return group;
        }

        static TermSet GetTermSet(Group group, string termSetName)
        {
            if (group == null) return null;


            var set = group.TermSets.FirstOrDefault(s => s.Name.Equals(termSetName, StringComparison.OrdinalIgnoreCase));
            if (set == null)
            {
                Console.BackgroundColor = System.ConsoleColor.Red;
                Console.WriteLine(
                    String.Format("The term set {0} could be not be found.", termSetName)
                );
                Console.ResetColor();
                
            }

            return set;

        }

    }
}


/*
using (SPSite site = new SPSite(suppliedParams["server"]))
{
    using (SPWeb web = site.OpenWeb())
    {
        TaxonomySession taxonomySession = new TaxonomySession(site);
        TermStore termStore = taxonomySession.TermStores["Metadata Service Application Proxy"];
        Group group = termStore.Groups[suppliedParams["group"]];
        TermSet termSet = group.TermSets[suppliedParams["termSet"]];
        Term term = null;
        if (suppliedParams.Keys.Contains("term")) term = termSet.Terms[suppliedParams["term"]];
        SPList list = web.Lists.TryGetList("Test");
        if (list != null)
        {
            TaxonomyField taxonomyField = list.Fields[suppliedParams["field"]] as TaxonomyField;
            TaxonomyFieldValue taxonomyFieldValue = new TaxonomyFieldValue(taxonomyField);
            taxonomyFieldValue.TermGuid = term.Id.ToString();
            taxonomyFieldValue.Label = term.Name;
            SPListItem item = list.Items.Add();
            item["Title"] = "Sample";
            item["TaxonomyField"] = taxonomyFieldValue;
            item.Update();
            list.Update();
        }
        else
        {
            Console.WriteLine(list.Title + " does not exists in the list");
        }
        Console.ReadLine();
    }
}

*/