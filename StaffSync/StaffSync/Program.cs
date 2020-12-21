using System;
using System.Linq;
using Microsoft.SharePoint;
using System;
using System.Web.Configuration;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Publishing;

namespace StaffSync
{
    using System.Collections;
    using Chalmers.Core.Common.Logging;
    using Chalmers.PublicWeb.Core;
    using Microsoft.SharePoint.Taxonomy;
    using System.Collections.Generic;
    using System.Linq;
    using System.Globalization;
    using Chalmers.Core;
    using Chalmers.Core.Configuration;
    using Chalmers.PublicWeb.Jobs; 
    
    
    class Program
    {
        static void Main(string[] args)
        {
            Uri ruri = new System.Uri("http://localhost:51001/");
            Console.WriteLine("Hämtar webbapplikationen...");
            SPWebApplication spwa = SPWebApplication.Lookup(ruri);
            string contentDbGuid = "6c1e7e35-9dfe-4974-955f-cd7e466e2432";
            Console.WriteLine("Skapar jobb...");
            ChalmersStaffSync css = new ChalmersStaffSync("Chalmers Staff Sync (Manual 1.0)", spwa);
            Console.WriteLine("Startar jobb...");
            css.Execute(new System.Guid(contentDbGuid));
        }
    }



    internal class ChalmersStaffSync : SPJobDefinition
    {
        private const string StaffContentTypeid =
            "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE8110054188dba4672400d8b90b9dd45ba5604";

        public ChalmersStaffSync() { }

        public ChalmersStaffSync(string jobName, SPWebApplication webApplication)
            : base(jobName, webApplication, null, SPJobLockType.ContentDatabase)
        {
            Title = "Chalmers Staff Sync";
        }



        public override void Execute(Guid contentDbId)
        {
                LoggingService.WriteTrace(EventSeverity.Information, "Starting Chalmers Staff Sync", LogCategory.ChalmersPublicWeb);

                SPWebApplication webApp;
                SPContentDatabase spContentDatabase;

                try
                {
                    webApp = Parent as SPWebApplication;
                    if (webApp == null) throw new ChalmersException("Web application does not exist");
                    spContentDatabase = webApp.ContentDatabases[contentDbId];
                }
                catch (Exception)
                {
                    throw new ChalmersException("Web application does not exist");
                }

                using (SPSite site = spContentDatabase.Sites[0])
                {
                    TermSet termSet = null;
                    Term departmentTerm = Chalmers.Core.Common.Utility.GetDepartmentsTerm(site, ref termSet);

                    if (departmentTerm == null || termSet == null)
                    {
                        LoggingService.WriteTrace(EventSeverity.Error, "TermStore unavailable. Unable to continue.", LogCategory.ChalmersPublicWeb);
                        return;
                    }

                    Dictionary<string, Guid> indexedTaxonomy = new Dictionary<string, Guid>();
                    Chalmers.Core.Common.Utility.BuildTaxonomyHashTable(departmentTerm, indexedTaxonomy, Chalmers.Core.Common.Utility.OrganizationDept.Alla);
                    
                    int itemsProcessed = 0;
                    List<string> staffPaths = new List<string>() {  "/en/staff", "/sv/personal", "/sv", "/en" };

                    foreach (string staffPath in staffPaths)
                    {
                        Console.WriteLine("Letar efter " + staffPath);
                        using (SPWeb spLanguageRoot = site.OpenWeb(staffPath))
                        {
                            List<Guid> stafWebId = new List<Guid>();

                            foreach (SPWeb currentWeb in spLanguageRoot.Webs)
                            {
                                try
                                {
                                    if (currentWeb.WebTemplate.ToUpper(CultureInfo.InvariantCulture) == "CHALMERSPUBLICWEB" && currentWeb.Configuration == 13)
                                    {
                                        stafWebId.Add(currentWeb.ID);
                                        LoggingService.WriteTrace(EventSeverity.Information, "ListUrl = " + currentWeb.ID, LogCategory.ChalmersPublicWeb);
                                    }
                                }
                                finally { if (currentWeb != null) currentWeb.Dispose(); } }

                            foreach (Guid id in stafWebId)
                            {
                                Console.WriteLine("Påbörjar bearbetning av " + staffPath);
                                LoggingService.WriteTrace(EventSeverity.Information, "Geting Staff User Detail", LogCategory.ChalmersPublicWeb);
                                SyncStaffList(site, id, indexedTaxonomy, termSet, ref itemsProcessed);
                            }
                        }
                    }
                }

                LoggingService.WriteTrace(EventSeverity.Information, "Ending Chalmers Staff Sync", LogCategory.ChalmersPublicWeb);
        }

        private void SyncStaffList(SPSite site, Guid stafWebId, Dictionary<string, Guid> indexedTaxonomy, TermSet termSet, ref int itemsProcessed)
        {
            string ldapServerUrl = ConfigurationFile.GetGonfigurationFile(Constants.ConfigurationFilePath).GetSettingValue("LdapServerUrl");
            
            if (string.IsNullOrEmpty(ldapServerUrl))
            {
                string msg = string.Format("No LDAP server configured! Unable to sync staff members. Add the key 'LdapServerUrl' to the configuration file {0}", Constants.ConfigurationFilePath);
                Console.WriteLine(msg);
                LoggingService.WriteTrace(
                    EventSeverity.Error,
                    msg, 
                    LogCategory.ChalmersPublicWeb);
                return;
            }

            using (SPWeb staffWeb = site.OpenWeb(stafWebId))
            {
                PublishingWeb pweb = PublishingWeb.GetPublishingWeb(staffWeb);
                SPList staffSitePagesList = pweb.PagesList;
                LoggingService.WriteTrace(EventSeverity.Information, "Got pages list: " + staffSitePagesList.Title, LogCategory.ChalmersPublicWeb);
                SPListItemCollection col = staffSitePagesList.Items;

                SPField field = (staffSitePagesList).Fields.GetField(Constants.ChalmersID);

                LoggingService.WriteTrace( EventSeverity.Information,
                    string.Format( "Got department taxonomy, number of items = {0}. Description with text 59fosmig exists = {1}", indexedTaxonomy.Count, indexedTaxonomy.ContainsKey("59fosmig")),
                    LogCategory.ChalmersPublicWeb);
                
                
                string principalCID = ConfigurationFile.GetGonfigurationFile(Chalmers.Core.Constants.CoreConstants.ConfigurationFilePath).GetSettingValue("PrincipalCID");

                foreach (SPListItem item in col)
                {
                    double max = 16000.0;
                    if (itemsProcessed < max)
                    {
                        try { UpdateProgress((int)((itemsProcessed++ / max) * 100)); } //col.Count 
                        catch { }
                        Console.Write("\rProcessing " + itemsProcessed + " " + item.Url + "        ");
                    }

                    try
                    {
                        if (item.ContentTypeId.ToString().StartsWith(StaffContentTypeid, StringComparison.OrdinalIgnoreCase))
                        {
                            LoggingService.WriteTrace( EventSeverity.Information, "Passed content type check.", LogCategory.ChalmersPublicWeb);

                            
                            var staffUserId = item.Fields[field.Id] as SPFieldUser;
                            if (staffUserId != null && item[Constants.ChalmersID] != null)
                            {
                                updateFromAD(indexedTaxonomy, termSet, ldapServerUrl, principalCID, item, staffUserId);
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        string message = string.Format("User Update Failed: {0} Details: {1}", item.Title, exception);
                        Console.WriteLine(Environment.NewLine + message);
                    }
                }
            }
        }

        private static void updateFromAD(Dictionary<string, Guid> indexedTaxonomy, TermSet termSet, string ldapServerUrl, string principalCID, SPListItem item, SPFieldUser staffUserId)
        {
            var chalmerUserId = staffUserId.GetFieldValue(item[Constants.ChalmersID].ToString()) as SPFieldUserValue;
            if (chalmerUserId != null)
            {
                string username = chalmerUserId.User.LoginName;
                LoggingService.WriteTrace(EventSeverity.Information, "Got username: " + username, LogCategory.ChalmersPublicWeb);
                string userLoginName = username;
                if (username.IndexOf('|') > 1)
                {
                    userLoginName = username.Split('|')[1];
                    LoggingService.WriteTrace(EventSeverity.Information, "UserLoginName has claims: " + userLoginName, LogCategory.ChalmersPublicWeb);
                }

                if (!string.IsNullOrEmpty(userLoginName))
                {
                    ADuser aduser = new ADuser(userLoginName) { 
                        LdapServerUrl = ldapServerUrl ,
                        GivenName = string.Empty,
                        SN = string.Empty,
                        Mail = string.Empty,
                        Organisation = string.Empty,
                        TelephoneNumber = string.Empty,
                        OtherTelephone = string.Empty
                    };


#if DEBUG
                    string value = ConfigurationFile.GetGonfigurationFile(Chalmers.PublicWeb.Core.Constants.ConfigurationFilePath).GetSettingValue("StaffSyncCID");
                    if (value.ToLower().Contains(aduser.CID.ToLower().Replace("net\\", string.Empty)))
                    {
                        string x = "asdf";
                        System.Diagnostics.Debug.Print(x.ToString());
                    }
#endif

                    LoggingService.WriteTrace(EventSeverity.Information, "aduser CID = " + aduser.CID, LogCategory.ChalmersPublicWeb);
                    string domain = WebConfigurationManager.AppSettings["CurrentDomain"];

                    if (string.IsNullOrEmpty(domain)) aduser.ADDomainName = @"net\";
                    else aduser.ADDomainName = domain + @"\";

                    if (aduser.CID.IndexOf(@"\", StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        aduser.CIDWithoutDomain = aduser.CID.Remove(0, aduser.CID.IndexOf(@"\", StringComparison.OrdinalIgnoreCase) + 1);
                        #if DEBUG
                        if (aduser.CIDWithoutDomain.ToString().ToLower().Equals("kain") ||
                            aduser.CIDWithoutDomain.ToString().ToLower().Equals("anderska"))
                        {
                            string x = aduser.CIDWithoutDomain.ToString();
                            System.Diagnostics.Debug.Print(x);
                        }
                        #endif
                    }
                    List<Term> terms = null;
                    TaxonomyField managedField = null;
                    aduser = ADUserProfile.GetUserProfileFromAD(aduser);

                    if (aduser != null)
                    {
                        bool adUnitsHaveChanged = taxFieldChanged(indexedTaxonomy, termSet, item, aduser, ref terms, ref managedField);
                        System.Text.StringBuilder phoneNumbers = new System.Text.StringBuilder();
                        if (!string.IsNullOrEmpty(aduser.TelephoneNumber)) phoneNumbers.Append(aduser.TelephoneNumber);
                        if (phoneNumbers.Length > 0 && !string.IsNullOrEmpty(aduser.OtherTelephone)) phoneNumbers.Append(", ");
                        if (!string.IsNullOrEmpty(aduser.OtherTelephone)) phoneNumbers.Append(aduser.OtherTelephone);

                        if (
                            hasUpdatedValues(item, aduser, phoneNumbers) ||
                            DateTime.Now.DayOfWeek == DayOfWeek.Saturday ||
                            adUnitsHaveChanged ||
                            (aduser.CID.Equals(principalCID))
                            )
                        {
                            item["EduPersonOrcid"]         = aduser.EduPersonOrcid;
                            item["FieldStaffGivenname"]    = aduser.GivenName;
                            item["FieldStaffLastName"]     = aduser.SN;
                            item["FieldStaffFullName"]     = aduser.GivenName + " " + aduser.SN;
                            item["Title"]                  = aduser.GivenName + " " + aduser.SN;
                            item["FieldStaffOrganisation"] = aduser.Organisation;
                            item["FieldStaffTelephone"]    = phoneNumbers.ToString();
                            item["OfficeRoomNumber"]       = aduser.OfficeRoomNumber;
                            item["OfficeStreet"]           = aduser.OfficeStreet;
                            item["OfficeFloorNumber"]      = aduser.OfficeFloorNumber;
                            
                            setEmail(principalCID, item, aduser);

                            if (adUnitsHaveChanged)
                            {
                                if (terms == null) terms = new List<Term>();
                                managedField.SetFieldValue(item, terms);
                            }
                            
                            trySaveChanges(item, aduser);
                        }
                    }
                }
            }
        }

        private static bool taxFieldChanged(
            Dictionary<string, Guid> indexedTaxonomy, 
            TermSet termSet, 
            SPListItem item, 
            ADuser aduser, 
            ref List<Term> terms, 
            ref TaxonomyField managedField)
        {
            managedField = item.Fields.GetFieldByInternalName("TaxChalmersDepartment") as TaxonomyField;
            if (aduser.AllUnits != null && aduser.AllUnits.Count > 0)
            {
                terms = aduser.AllUnits.Where(indexedTaxonomy.ContainsKey).Select(s => termSet.GetTerm(indexedTaxonomy[s])).ToList();
                if (terms.Count > 0)
                {
                    LoggingService.WriteTrace(EventSeverity.Information, string.Format("Found {0} matching terms", terms.Count), LogCategory.ChalmersPublicWeb);
                    
                    if (managedField != null)
                    {
                        string fieldTaxItems = item["TaxChalmersDepartment"].ToString();

                        for (int i = 0; i < terms.Count(); i++)
                        {
                            if (fieldTaxItems.Count(f => f == '|') != terms.Count() ||
                                !fieldTaxItems.Contains(terms[i].Name))
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            else {
                string fieldTaxItems = item["TaxChalmersDepartment"].ToString();
                if (fieldTaxItems.Count(f => f == '|') > 0) return true;
            }
            return false;
        }

        private static void trySaveChanges(SPListItem item, ADuser aduser)
        {
            if (item.File.Level != SPFileLevel.Checkout)
            {
                LoggingService.WriteTrace(
                    EventSeverity.Information,
                    string.Format("User {0} found and page is not checked out, proceeding to update page item", aduser.CID),
                    LogCategory.ChalmersPublicWeb);
                item.SystemUpdate(false);
            }
            else
            {
                string info = string.Format("Page item for user {0} is checked out and cannot be updated", aduser.CID);
                LoggingService.WriteTrace(
                    EventSeverity.Information,
                    info,
                    LogCategory.ChalmersPublicWeb);
                Console.WriteLine(Environment.NewLine + info);
            }
        }

        private static void setEmail(string principalCID, SPListItem item, ADuser aduser)
        {
            if (aduser.CID.Equals(principalCID))
            {
                string principalEMail = ConfigurationFile.GetGonfigurationFile(Chalmers.Core.Constants.CoreConstants.ConfigurationFilePath).GetSettingValue("PrincipalEMail");
                LoggingService.WriteTrace(
                    EventSeverity.Information,
                    "Principal CID found: '" + aduser.CID + "'. Setting e-mail to: '" + principalEMail + "', based on: '" + Chalmers.Core.Constants.CoreConstants.ConfigurationFilePath + "'.",
                    LogCategory.ChalmersPublicWeb);
                item["FieldStaffEmail"] = principalEMail;
            }
            else
            {
                item["FieldStaffEmail"] = aduser.Mail;
            }
        }


        private static bool hasUpdatedValues(SPListItem item, ADuser aduser, System.Text.StringBuilder phoneNumbers)
        {
            string FieldStaffGivenname = item["FieldStaffGivenname"] == null ? string.Empty : item["FieldStaffGivenname"].ToString();
            if (!FieldStaffGivenname.Equals(aduser.GivenName)) return true;

            string FieldStaffLastName = item["FieldStaffLastName"] == null ? string.Empty : item["FieldStaffLastName"].ToString();
            if (!FieldStaffLastName.Equals(aduser.SN)) return true;
            
            string FieldStaffFullName = item["FieldStaffFullName"] == null ? string.Empty : item["FieldStaffFullName"].ToString();
            if (!FieldStaffFullName.Equals(aduser.GivenName + " " + aduser.SN)) return true;

            string Title = item["Title"] == null ? string.Empty : item["Title"].ToString();
            if (!Title.Equals(aduser.GivenName + " " + aduser.SN)) return true;
            
            string FieldStaffOrganisation = item["FieldStaffOrganisation"] == null ? string.Empty : item["FieldStaffOrganisation"].ToString();
            if (!FieldStaffOrganisation.Equals(aduser.Organisation)) return true;
            
            string FieldStaffTelephone = item["FieldStaffTelephone"] == null ? string.Empty : item["FieldStaffTelephone"].ToString();
            if (!FieldStaffTelephone.Equals(phoneNumbers.ToString())) return true;

            string eduPersonOrcID = item["EduPersonOrcid"] == null ? string.Empty : item["EduPersonOrcid"].ToString();
            string aduserEduPersonOrcid = aduser.EduPersonOrcid == null ? string.Empty : aduser.EduPersonOrcid;
            if (!eduPersonOrcID.Equals(aduserEduPersonOrcid)) return true;

            string RoomNumber = item["OfficeRoomNumber"] == null ? string.Empty : item["OfficeRoomNumber"].ToString();
            if (!RoomNumber.Equals(aduser.OfficeRoomNumber.ToString())) return true;

            string Street = item["OfficeStreet"] == null ? string.Empty : item["OfficeStreet"].ToString();
            if (!Street.Equals(aduser.OfficeStreet.ToString())) return true;

            // Since we currently guess floor number we don't want to overwrite manual changes so we only set this value if there is no value.
            if (item["OfficeFloorNumber"] == null) return true; 
            //string OfficeFloorNumber = item["OfficeFloorNumber"] == null ? string.Empty : item["OfficeFloorNumber"].ToString();
            //if (!OfficeFloorNumber.Equals(aduser.OfficeFloorNumber.ToString())) return true;


            return false;
        }




    }

}
