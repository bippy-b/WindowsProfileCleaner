using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//User added
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Configuration;




namespace WindowsProfileCleaner
{
    public static class Library

    {
        //[DllImport("userenv.dll", CharSet = CharSet.Unicode, ExactSpelling = false, SetLastError = true)]
        //public static extern bool DeleteProfile(string sidString, string profilePath, string ComputerName);

        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        //public static void WriteErrorLog(Exception ex, String step)
        //{
        //    StreamWriter sw = null;
        //    try
        //    {
        //        sw = new StreamWriter(Properties.Settings.Default.LogPath + "WindowsProfileCleaner_" + DateTime.Now.ToString("yyyyMMdd") + ".log", true);
        //        sw.WriteLine(DateTime.Now.ToString() + " - Step " + step + ": " + ex.Source.ToString() + ": " + ex.Message.ToString());
        //        sw.Flush();
        //        sw.Close();
        //        //sw.Dispose();
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.Message);
        //        //Need to add writing to the application log
        //    }

        //}
        //public static void WriteErrorLog(String message)
        //{
        //    StreamWriter sw = null;
        //    try
        //    {
        //        sw = new StreamWriter(Properties.Settings.Default.LogPath + "WindowsProfileCleaner_" + DateTime.Now.ToString("yyyyMMdd") + ".log", true);
        //        sw.WriteLine(DateTime.Now.ToString() + " - " + message);
        //        sw.Flush();
        //        sw.Close();
        //        //sw.Dispose();
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.Message);
        //        //Need to add writing to the application log
        //    }

        //}

        public static void WriteBatchFile(String message, bool Overwrite)
        {
            StreamWriter sw = null;

            logger.Debug("WriteBatchFile:" + message);
            //if (Properties.Settings.Default.Debug == "true" || Properties.Settings.Default.Debug == "all")
            //{
            //    Library.WriteErrorLog("WriteBatchFile:" + message);
            //}

            try
            {
                if (Overwrite == true)
                {
                    sw = new StreamWriter(Properties.Settings.Default.BATPath + "WindowsProfileCleaner.bat", false);
                }
                else
                {
                    sw = new StreamWriter(Properties.Settings.Default.BATPath + "WindowsProfileCleaner.bat", true);
                }
                sw.WriteLine(message);
                sw.Flush();
                sw.Close();
                //sw.Dispose();
            }
            catch (Exception e)
            {
                logger.Error(e.Message, "Error occured in WriteBatchFile");
                //Need to add writing to the application log
            }

        }

        public static UserPrincipal LookupUserWithDomain(String AcctName, String UserDomain)
        {
            UserPrincipal UserAcct = null;

            try
            {
                // enter AD settings  
                PrincipalContext AD = new PrincipalContext(ContextType.Domain, UserDomain);

                // create search user and add criteria  
                //Console.Write("Enter logon name: ");
                UserPrincipal u = new UserPrincipal(AD);
                u.SamAccountName = AcctName;

                // search for user  
                PrincipalSearcher search = new PrincipalSearcher(u);
                UserPrincipal result = (UserPrincipal)search.FindOne();
                search.Dispose();

                if (result != null)
                {
                    //if (Properties.Settings.Default.Debug == "true" || Properties.Settings.Default.Debug == "all")
                    //{
                    //    Library.WriteErrorLog("LookupUser DisplayName:" + result.DisplayName);
                    //    Library.WriteErrorLog("LookupUser LastLogon:" + result.LastLogon);
                    //}
                    logger.Debug("LookupUser DisplayName:" + result.DisplayName);
                    logger.Debug("LookupUser LastLogon:" + result.LastLogon);
                    return result;
                }

            }
            catch (Exception e)
            {
                logger.Error(e.Message, "LookupUser error:");
            }

            return UserAcct;

        }

        public static UserPrincipal LookupUser(String AcctName, String UserSid)
        {

            string tree = Properties.Settings.Default.SearchDomains;
            string[] domains = tree.Split(',');

            //if (Properties.Settings.Default.Debug == "true" || Properties.Settings.Default.Debug == "all")
            //{
            //    WriteErrorLog("Lookup Account:" + AcctName);
            //    WriteErrorLog("Lookup Account Sid:" + UserSid);
            //}
            logger.Debug("Lookup Account:" + AcctName);
            logger.Debug("Lookup Account Sid:" + UserSid);

            //loop through the comma delimited domain list
            foreach (string domain in domains)
            {
                try
                {
                    //if (Properties.Settings.Default.Debug == "true" || Properties.Settings.Default.Debug == "all")
                    //{
                    //    WriteErrorLog("Lookup Account in domain:" + domain);
                    //}
                    logger.Debug("Lookup Account in domain:" + domain);

                    // enter AD settings  
                    PrincipalContext AD = new PrincipalContext(ContextType.Domain, domain);

                    // create search user and add criteria  
                    UserPrincipal u = new UserPrincipal(AD);
                    u.SamAccountName = AcctName;

                    // search for user  
                    PrincipalSearcher search = new PrincipalSearcher(u);
                    UserPrincipal result = (UserPrincipal)search.FindOne();
                    search.Dispose();

                    if (result != null)
                    {
                        if (result.Sid.ToString() == UserSid)
                        {
                            //UserDomain = "iconcr.com";
                            //Library.WriteErrorLog(domain + ":Account found - User DisplayName:" + result.DisplayName);
                            logger.Info(domain + ":Account found - User DisplayName:" + result.DisplayName);
                            return result;
                        }
                        else
                        {
                            //Library.WriteErrorLog(AcctName + "/" + domain + ":An account matching this users account was found but the sids did not match");

                            logger.Info(AcctName + "/" + domain + ":An account matching this users account was found but the sids did not match");
                            
                            //if (Properties.Settings.Default.Debug == "true" || Properties.Settings.Default.Debug == "all")
                            //{
                            //    WriteErrorLog("Found Account Sid:" + result.Sid.ToString());
                            //    WriteErrorLog("Account we are searching for Sid:" + UserSid);
                            //}

                            logger.Debug("Found Account Sid:" + result.Sid.ToString());
                            logger.Debug("Account we are searching for Sid:" + UserSid);
                        }
                    }

                }
                catch (Exception e)
                {
                    //Library.WriteErrorLog(domain + ":Error during LookupDomain function:" + e.Message);
                    logger.Error(e.Message, domain + ":Error during LookupDomain function:");
                }
            }


            return null;

        }
        public static Boolean VerifyAccount(String AcctName)
        {

            bool AcctIsValid = true;
            string val = Properties.Settings.Default.AccountIgnore;
            string[] words = val.Split(',');

            //if (Properties.Settings.Default.Debug == "true" || Properties.Settings.Default.Debug == "all")
            //{
            //    WriteErrorLog("Verifying Account:" + AcctName);
            //    if (Properties.Settings.Default.Debug == "all")
            //    {
            //        WriteErrorLog("Verification string length: " + val.Length);
            //        WriteErrorLog("Verification string count:" + words.Count());
            //    }
            //}
            logger.Debug("Verifying Account:" + AcctName);
            logger.Trace("Verification string length: " + val.Length);
            logger.Trace("Verification string count:" + words.Count());
            
            //Loop through the accounts to ignore
            foreach (string word in words)
            {

                //if (Properties.Settings.Default.Debug == "all")
                //{
                //    WriteErrorLog("Verifying Account:word:foreachloop:" + word);
                //    WriteErrorLog("Verifying Account:AcctIsValid:" + AcctIsValid.ToString());
                //}

                logger.Trace("Verifying Account:word:foreachloop:" + word);
                logger.Trace("Verifying Account:AcctIsValid:" + AcctIsValid.ToString());

                //If AcctIsValid is false no need to continue checking.  Just move on.
                if (AcctIsValid == true)
                {
                    if (AcctName.ToLower().Equals(word.ToLower()))
                    {
                        string msg = "The account " + AcctName + " contains an excluded account name:" + word;
                        //WriteErrorLog(msg);
                        logger.Info(msg);
                        AcctIsValid = false;
                        
                    }
                }

            }
                        
            return AcctIsValid;
        }

        public static List<string> GetLocalComputerUsers()
        {
            // create your domain context
            // search your domain
            //PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            // search your local machine
            PrincipalContext ctx = new PrincipalContext(ContextType.Machine);

            // define a "query-by-example" principal - here, we search for a UserPrincipal 
            UserPrincipal qbeUser = new UserPrincipal(ctx);

            // create your principal searcher passing in the QBE principal    
            PrincipalSearcher srch = new PrincipalSearcher(qbeUser);

            List<string> userNames = new List<string>();

            //Library.WriteErrorLog("List<string> userNames");
            logger.Info("List<string> userNames");

            // find all matches
            foreach (var found in srch.FindAll())
            {
                //Library.WriteErrorLog("Name  :" + found.Name.ToString());
                logger.Info("Name  :" + found.Name.ToString());
                // do whatever here - "found" is of type "Principal"
                if (!userNames.Contains(found.Name))
                {
                    userNames.Add(found.Name);
                }
            }
            return userNames;
        }

        public static List<string> GetDomainUsers(String domain)
        {
            // create your domain context
            // search your domain
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain,domain);

            // search your local machine
            //PrincipalContext ctx = new PrincipalContext(ContextType.Machine);

            // define a "query-by-example" principal - here, we search for a UserPrincipal 
            UserPrincipal qbeUser = new UserPrincipal(ctx);

            // create your principal searcher passing in the QBE principal    
            PrincipalSearcher srch = new PrincipalSearcher(qbeUser);

            List<string> userNames = new List<string>();

            //Library.WriteErrorLog("List<string> userNames");
            logger.Info("List<string> userNames");

            // find all matches
            foreach (var found in srch.FindAll())
            {
                logger.Info("Name  :" + found.Name.ToString());
                //Library.WriteErrorLog("Name  :" + found.Name.ToString());
                // do whatever here - "found" is of type "Principal"
                if (!userNames.Contains(found.Name))
                {
                    userNames.Add(found.Name);
                }
            }
            return userNames;
        }
    }
}
