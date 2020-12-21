namespace Chalmers.PublicWeb.Jobs
{

    using System;
    using System.Collections.Generic;
    using System.DirectoryServices;
    using System.Linq;
    using Chalmers.Core.Common.Logging;
    using Microsoft.SharePoint.Administration;
    using System.Globalization;
    using Chalmers.Core;

    public class ADuser
    {
        public const string Profile_FLD_GivenName = "FirstName";
        public const string Profile_FLD_LASTNAME = "LastName";
        public const string Profile_FLD_MAIL = "WorkEmail";
        public const string Profile_FLD_TELEPHONENUMBER = "WorkPhone";
        public const string Profile_FLD_PREFERDNAME = "PreferredName";
        public const string Profile_FLD_ORGANISATION = "Department";

        public ADuser(string cid)
        {
            this.CID = cid;
        }

        public string GivenName { get; set; }
        public string SN { get; set; }
        public string Mail { get; set; }
        public string TelephoneNumber { get; set; }
        public string OtherTelephone { get; set; }
        public string EduPersonOrgUnitDN { get; set; }
        public List<string> AllUnits { get; set; }
        public string Organisation { get; set; }
        public string CID { get; set; }
        public string CIDWithoutDomain { get; set; }
        public string LdapServerUrl { get; set; }
        public string ADDomainName { get; set; }
        public string EduPersonOrcid { get; set; }
        public string OfficeRoomNumber { get; set; }
        public string OfficeStreet { get; set; }
        public string OfficeFloorNumber { get; set; }
        
    }

    public static class ADUserProfile
    {
        public static ADuser GetUserProfileFromAD(ADuser aduser)
        {
            if (aduser == null) throw new ArgumentNullException("aduser");
            DirectoryEntry entry = null;

            try
            {
                string message = String.Format("Attempting to query the AD for user {0}", aduser.CID);
                LoggingService.WriteTrace(EventSeverity.Information, message, LogCategory.ChalmersPublicWeb);

                string filterAttribute = string.Empty, otherTelephone = string.Empty;
                int mode = 0; // 0 = ldap.chalmers.se, 1 = net.chalmers.se

                // HACK: Directory search settings based on URL
                if (aduser.LdapServerUrl.ToUpperInvariant().Contains("LDAP.CHALMERS.SE"))
                {
                    entry = new DirectoryEntry(aduser.LdapServerUrl, string.Empty, string.Empty, AuthenticationTypes.Anonymous);
                    filterAttribute = "uid";
                }
                else
                {
                    entry = new DirectoryEntry(aduser.LdapServerUrl);
                    filterAttribute = "cn";
                    mode = 1;
                }

                message = String.Format("Created new DirectoryEntry for URL {0}", aduser.LdapServerUrl);
                LoggingService.WriteTrace(EventSeverity.Information, message, LogCategory.ChalmersPublicWeb);

                using (DirectorySearcher mainSearch = new DirectorySearcher(entry))
                {
                    mainSearch.Filter = string.Format("{0}={1}", filterAttribute, aduser.CIDWithoutDomain);
                    message = String.Format("Created new DirectorySearcher with filter {0}", mainSearch.Filter);
                    LoggingService.WriteTrace(EventSeverity.Information, message, LogCategory.ChalmersPublicWeb);

                    mainSearch.PropertiesToLoad.Add("givenName");
                    mainSearch.PropertiesToLoad.Add("sn");
                    mainSearch.PropertiesToLoad.Add("mail");
                    mainSearch.PropertiesToLoad.Add("telephoneNumber");
                    if (mode == 1) mainSearch.PropertiesToLoad.Add("otherTelephone");
                    mainSearch.PropertiesToLoad.Add("eduPersonOrgUnitDN");
                    mainSearch.PropertiesToLoad.Add("eduPersonOrcid");
                    mainSearch.PropertiesToLoad.Add("roomNumber");
                    mainSearch.PropertiesToLoad.Add("street");

                    SearchResult mainResult = mainSearch.FindOne();
                    LoggingService.WriteTrace(EventSeverity.Information, "AD search initialized", LogCategory.ChalmersPublicWeb);

                    LoggingService.WriteTrace(EventSeverity.Information, "AD search initialized", LogCategory.ChalmersPublicWeb);

                    if (mainResult != null)
                    {
                        LoggingService.WriteTrace(EventSeverity.Information, "AD search successful, creating user object.", LogCategory.ChalmersPublicWeb);
                        
                        foreach (string key in mainResult.Properties.PropertyNames)
                        {
                            string myKey = key.ToUpperInvariant();
                            int addedTelePhoneNumberes = 0;
                            foreach (string propValue in mainResult.Properties[key])
                            {
                                if (myKey.Equals("GIVENNAME")) aduser.GivenName = propValue;
                                else if (myKey.Equals("SN")) aduser.SN = propValue;
                                else if (myKey.Equals("MAIL")) aduser.Mail = propValue;
                                else if (myKey.Equals("TELEPHONENUMBER"))
                                {
                                    if (addedTelePhoneNumberes == 0) aduser.TelephoneNumber = propValue;
                                    else aduser.OtherTelephone = propValue;
                                    addedTelePhoneNumberes++;
                                }
                                else if (myKey.Equals("OTHERTELEPHONE")) aduser.OtherTelephone = propValue;
                                else if (myKey.Equals("EDUPERSONORCID")) aduser.EduPersonOrcid = propValue;
                                else if (myKey.Equals("ROOMNUMBER"))
                                {
                                    aduser.OfficeRoomNumber = propValue == null ? string.Empty : propValue;

                                    // Since we don't get floor number yet we guess based on known rules.
                                    for (int i = 0; i < propValue.Length; i++)
                                    {
                                        int outVal = 0;
                                        aduser.OfficeFloorNumber = propValue.Substring(i, 1);
                                        if (int.TryParse(aduser.OfficeFloorNumber, out outVal))
                                        {
                                            break;
                                        }
                                    }
                                    if (aduser.OfficeFloorNumber == null) aduser.OfficeFloorNumber = string.Empty;
                                    
                                }
                                else if (myKey.Equals("STREET")) aduser.OfficeStreet = propValue == null ? string.Empty : propValue;
                                else if (myKey.Equals("EDUPERSONORGUNITDN"))
                                {
                                    aduser.EduPersonOrgUnitDN = propValue;
                                    var allOrgProperties = mainResult.Properties[key].Cast<object>().Aggregate(string.Empty, (current, t) => string.IsNullOrEmpty(t.ToString()) ? current : string.Format("{0},{1}", current, t.ToString()));
                                    aduser.AllUnits = System.Text.RegularExpressions.Regex.Replace(allOrgProperties, "(OU|DC)=", string.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase).Split(
                                                        new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Distinct().ToList();
                                }

                            }
                        }

                        if (!string.IsNullOrEmpty(aduser.EduPersonOrgUnitDN))
                        {
                            // HACK: Directory search settings based on URL
                            using (DirectoryEntry entry2 = aduser.LdapServerUrl.ToUpperInvariant().Contains("LDAP.CHALMERS.SE")
                                ? new DirectoryEntry(aduser.LdapServerUrl + "/" + aduser.EduPersonOrgUnitDN, string.Empty, string.Empty, AuthenticationTypes.Anonymous)
                                : new DirectoryEntry(aduser.LdapServerUrl + "/" + aduser.EduPersonOrgUnitDN))
                            {
                                foreach (string key in entry2.Properties.PropertyNames)
                                {
                                    foreach (Object propValue in entry2.Properties[key])
                                    {
                                        string myKey = key.ToUpperInvariant();
                                        string value = propValue as string;

                                        if (myKey.Equals("DESCRIPTION"))
                                        {
                                            aduser.Organisation = value;
                                        }
                                    }
                                }
                            }
                        }

                        message =
                            string.Format(
                                "AD user info - CID:{0}, GivenName:{1}, SurName:{2}, E-mail:{3}, Organisation:{4}, Phone number:{5}, Phone number2:{6}, OrcId:{7}",
                                aduser.CID, aduser.GivenName, aduser.SN, aduser.Mail, aduser.Organisation, aduser.TelephoneNumber, aduser.OtherTelephone, aduser.EduPersonOrcid);
                        LoggingService.WriteTrace(EventSeverity.Information, message, LogCategory.ChalmersPublicWeb);

                        if (aduser.OfficeRoomNumber == null) aduser.OfficeRoomNumber = string.Empty;
                        if (aduser.OfficeFloorNumber == null) aduser.OfficeFloorNumber = string.Empty;
                        if (aduser.OfficeStreet == null) aduser.OfficeStreet = string.Empty;

                        return aduser;
                    }
                }
                return null;
            }
            catch (Exception exception)
            {
                throw new ChalmersException("Failed to bind: " + exception.Message, exception);
            }
            finally
            {
                if (entry != null)
                {
                    entry.Dispose();
                }
            }
        }

        private static void getOrganization(ADuser aduser)
        {
            if (!string.IsNullOrEmpty(aduser.EduPersonOrgUnitDN))
            {
                // HACK: Directory search settings based on URL
                using (DirectoryEntry entry2 = aduser.LdapServerUrl.ToUpperInvariant().Contains("LDAP.CHALMERS.SE")
                    ? new DirectoryEntry(aduser.LdapServerUrl + "/" + aduser.EduPersonOrgUnitDN, string.Empty, string.Empty, AuthenticationTypes.Anonymous)
                    : new DirectoryEntry(aduser.LdapServerUrl + "/" + aduser.EduPersonOrgUnitDN))
                {
                    foreach (string key in entry2.Properties.PropertyNames)
                    {
                        string myKey = key.ToUpperInvariant();
                        if (myKey.Equals("DESCRIPTION"))
                        {
                            foreach (Object propValue in entry2.Properties[key])
                            {
                                string value = propValue as string;
                                aduser.Organisation = value;
                                return;
                            }
                        }
                    }
                }
            }
        }
    }
}
