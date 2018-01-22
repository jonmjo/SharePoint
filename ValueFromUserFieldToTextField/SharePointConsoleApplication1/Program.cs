using System;
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.SharePoint.Taxonomy;
using Microsoft.SharePoint.Publishing;
using System.Net.Mail;
using System.Text;

namespace SharePointConsoleApplication1
{
    class Program
    {
        private const string fieldNameFrom = "CID";
        private const string fieldNameTo = "PDBPersonID";
        private static StringBuilder runLog = new StringBuilder();

        static void Main(string[] args)
        {
            if (args.Length == 1) StartProcess(args[0].ToString());
            else PrintHelpText();
        }


        private static void StartProcess(string url)
        {

            try
            {
                Console.Write("\nOpening '" + url + "'...");

                using (SPSite oSite = new SPSite(url))
                {
                    string[] webUrls = new string[] { "/en/staff/", "/en/staff/edit/", "/sv/personal/", "/sv/personal/redigera/"};

                    foreach (string webUrl in webUrls)
                    {
                        Console.Write("\nOpening '" + webUrl + "'...");
                        using (SPWeb oWeb = oSite.OpenWeb(webUrl))
                        {
                            PublishingWeb pweb = PublishingWeb.GetPublishingWeb(oWeb);
                            SPList staffSitePagesList = pweb.PagesList;
                            SPListItemCollection col = staffSitePagesList.Items;

                            SPField fieldFrom = (staffSitePagesList).Fields.GetField(fieldNameFrom);
                            //SPField fieldTo = (staffSitePagesList).Fields.GetField(fieldNameTo);

                            log("\nStarting processing pages...");
                            int i = 0;
                            foreach (SPListItem item in col)
                            {
                                log(i++ + " Doing: " + item.Url);

                                SPFieldUser staffUserId = item.Fields[fieldFrom.Id] as SPFieldUser;
                                if (item[fieldNameFrom] == null)
                                {
                                    log("  Item does not contain CID-value.");
                                    continue;
                                }

                                var chalmerUserId = staffUserId.GetFieldValue(item[fieldNameFrom].ToString()) as SPFieldUserValue;
                                if (item[fieldNameTo] == null) log("  To field is null");
                                else log("  To field is: " + item[fieldNameTo].ToString());

                                string username = chalmerUserId.User.LoginName.ToLower().Replace("net\\", string.Empty);
                                log("  Writing value: " + username);
                                item[fieldNameTo] = username;
                                log("  To field value is now: " + item[fieldNameTo].ToString());
                                if (item.File.Level != SPFileLevel.Checkout)
                                {
                                    log("  Saving...");
                                    item.SystemUpdate(false);
                                    log("  Done.");
                                }
                                else
                                {
                                    log("  Checked out file. Can't update.");
                                }

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("\n\nAn Error occured " + ex.Message + "\n\nEnsure that you have Administrative rights.\n"); // Also ensure platform target is x64 and not x86.
                Console.ForegroundColor = ConsoleColor.Gray;
                PrintHelpText();
            }

            sendEmail("ix3mjjo@chalmers.se", "Logg value from user field to text field", runLog.ToString());
        }

        private static void log(string msg)
        {
            runLog.Append(msg + Environment.NewLine);
            Console.WriteLine(msg);
        }

        private static void PrintHelpText()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            log("This program will copy the value of the CID field to the PDBPersonID-field.");
            log(
                string.Format(
                    "Start program by: {0} {1}",
                    System.AppDomain.CurrentDomain.FriendlyName,
                    "http://localhost:51001/"
                )
            );
            Console.ForegroundColor = ConsoleColor.Gray;
        }

        public static bool sendEmail(string emailaddress, string subject, string message)
        {
            try
            {
                if (string.IsNullOrEmpty(emailaddress)) throw new Exception("No e-mail address specified.");

                MailMessage mail = new MailMessage("noreply@ita.chalmers.se", emailaddress);
                mail.Subject = subject;
                mail.Body = message;

                SmtpClient client = new SmtpClient();
                client.Port = 25;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Host = "smtp.chalmers.se";
                client.Send(mail);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Kunde inte skicka e-post: " + ex.Message);
                return false;
            }
        }


    } // class
} // namespace
