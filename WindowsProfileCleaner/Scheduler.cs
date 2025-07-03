using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

//User Defined
using System.Timers;
using System.Configuration;
using System.Management;
using System.DirectoryServices.AccountManagement;
using Microsoft.Win32;


namespace WindowsProfileCleaner
{


    public partial class Scheduler : ServiceBase
    {
        private System.Timers.Timer timer1 = new Timer();

        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        public Scheduler()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            logger.Info("WindowsProfileCleaner Service ==Begin Start Up==");
            //Library.WriteErrorLog("WindowsProfileCleaner Service ==Begin Start Up==");
            //var appSettings = ConfigurationManager.AppSettings;
            //need to update it to read from the configuration
            //timer1.Interval = 30000;  //interval divided by 1000 = number of seconds

            //Initial interval setting very low so it fires quickly
            timer1.Interval = 5000;

            //timer1.Interval = TryParse.Int(System.Configuration.ConfigurationManager.AppSettings["Interval"]);
            timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Tick);
            
            //Library.WriteErrorLog("WindowsProfileCleaner Service Started.");
            //Library.WriteErrorLog("========Begin Settings Values=========");
            //Library.WriteErrorLog("Timer Interval:" + Properties.Settings.Default.CheckInterval);
            //Library.WriteErrorLog("LogPath:" + Properties.Settings.Default.LogPath);
            //Library.WriteErrorLog("Debug:" + Properties.Settings.Default.Debug);
            //Library.WriteErrorLog("AccountIgnore:" + Properties.Settings.Default.AccountIgnore);
            //Library.WriteErrorLog("DelProfile:" + Properties.Settings.Default.DelProfile);
            //Library.WriteErrorLog("SearchDomains:" + Properties.Settings.Default.SearchDomains);
            //Library.WriteErrorLog("EXEPath:" + Properties.Settings.Default.EXEPath);
            //Library.WriteErrorLog("BATPath:" + Properties.Settings.Default.BATPath);
            //Library.WriteErrorLog("========End Settings Values=========");

            logger.Info("WindowsProfileCleaner Service Started.");
            logger.Info("========Begin Settings Values=========");
            logger.Info("Timer Interval:" + Properties.Settings.Default.CheckInterval);
            logger.Info("LogPath:" + Properties.Settings.Default.LogPath);
            logger.Info("Debug:" + Properties.Settings.Default.Debug);
            logger.Info("AccountIgnore:" + Properties.Settings.Default.AccountIgnore);
            logger.Info("DelProfile:" + Properties.Settings.Default.DelProfile);
            logger.Info("SearchDomains:" + Properties.Settings.Default.SearchDomains);
            logger.Info("EXEPath:" + Properties.Settings.Default.EXEPath);
            logger.Info("BATPath:" + Properties.Settings.Default.BATPath);
            logger.Info("========End Settings Values=========");

            timer1.Start();

            logger.Trace("timer1 started");
        }

        protected override void OnStop()
        {
            logger.Info("WindowsProfileCleaner Service Stopped.");
            //Library.WriteErrorLog("WindowsProfileCleaner Service Stopped.");
            
        }

        private void timer1_Tick(object sender, ElapsedEventArgs e)
        {
            bool FirstUser = true;

            try
            {
                logger.Info("WindowsProfileCleaner Service timer has elapsed.");
                //Library.WriteErrorLog("WindowsProfileCleaner Service timer has elapsed.");

               //Set the timer to the correct value after the initial trigger
               if (timer1.Interval < 10000)
               {
                    Double myInterval;
                    if (Double.TryParse(Properties.Settings.Default.CheckInterval, out myInterval))
                    {
                        timer1.Interval = System.Convert.ToDouble(Properties.Settings.Default.CheckInterval);
                    }
                    else
                    {
                        timer1.Interval = 60000;
                    }
               }

               RegistryKey rootKey = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList", false);

                string[] keys = rootKey.GetSubKeyNames();

                //Log the number of profiles found.  This way we can see progress.
                //Library.WriteErrorLog("Number of profiles in ProfileList:" + keys.Count());
                logger.Info("Number of profiles in ProfileList:" + keys.Count());

                logger.Debug("rootKey:keys: Number of keys in ProfileList:" + keys.Count());

                //if (Properties.Settings.Default.Debug == "true")
                //{
                //    Library.WriteErrorLog("rootKey:keys: Number of keys in ProfileList:" + keys.Count());
                //}
                

                logger.Trace("FirstUser:" + FirstUser.ToString());

            foreach (string key in keys)
            {
                RegistryKey newkey = rootKey.OpenSubKey(key);

                    logger.Debug("foreachkeyloop:" + newkey.ToString());

                    //if (Properties.Settings.Default.Debug == "true")
                    //if (Properties.Settings.Default.Debug == "true" )
                    //{
                    //    Library.WriteErrorLog("foreachkeyloop:" + newkey.ToString());
                    //}

                

                //if (key.ToString().StartsWith("S-1-5-21"))
                //Only user created logins start with S-1-5-21.  This way we avoid some special accounts.
                if (newkey.Name.ToString().Replace("HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList\\", "").StartsWith("S-1-5-21"))
                //if (newkey.GetValue("ProfileImagePath").ToString().StartsWith("C:\\Users"))
                {
                    //get the account name
                    string UserAccount = newkey.GetValue("ProfileImagePath").ToString().Replace("C:\\Users\\", "").ToString();
                    string UserDomain = "";
                    string UserSid = newkey.Name.ToString().Replace("HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList\\", "");

                    //trace data
                    logger.Trace("UserAccount:"+ UserAccount);
                    logger.Trace("UserSid:"+ UserSid);

                    if (UserAccount.Contains("."))
                    {
                        UserDomain = UserAccount.Split('.').Last();
                        UserAccount = UserAccount.Split('.').First();

                        //User Domain was found
                        logger.Trace("Users domain is:" + UserDomain);

                        //if (Properties.Settings.Default.Debug == "true" || Properties.Settings.Default.Debug == "all")
                        //{
                        //    Library.WriteErrorLog("Users domain is:" + UserDomain);
                        //}
                        

                    }

                        logger.Trace("---===Begin Key===---");
                        logger.Trace(newkey.Name.ToString().Replace("HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList\\", ""));
                        logger.Trace(newkey.GetValue("ProfileImagePath").ToString().Replace("C:\\Users\\", ""));
                        logger.Trace("---===End Key===---");

                    //if (Properties.Settings.Default.Debug == "all")
                    //{
                    //    Library.WriteErrorLog("---===Begin Key===---");
                    //    Library.WriteErrorLog(newkey.Name.ToString().Replace("HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList\\", ""));
                    //    Library.WriteErrorLog(newkey.GetValue("ProfileImagePath").ToString().Replace("C:\\Users\\", ""));
                    //    //Library.WriteErrorLog("Validation:" + Library.VerifyAccount(newkey.GetValue("ProfileImagePath").ToString().Replace("C:\\Users\\", "")).ToString());
                    //    Library.WriteErrorLog("---===End Key===---");
                    //}

                    //VerifyAccount will return true if it is a valid account to process
                    if (Library.VerifyAccount(UserAccount) == true)
                    {
                        UserPrincipal UserAcct = Library.LookupUser(UserAccount, UserSid);
                        
                        //Check if an account was found ( not null )
                        if (UserAcct != null)
                        {
                                logger.Info("For user account:" + UserAccount + " an account was found.  The account enabled is: " + UserAcct.Enabled.ToString());
                                //if (Properties.Settings.Default.Debug == "true" || Properties.Settings.Default.Debug == "all")
                                //{
                                //    Library.WriteErrorLog("For user account:" + UserAccount + " an account was found.  The account enabled is: " + UserAcct.Enabled.ToString());
                                //}

                            //if the account enabled = false then delete the profile
                            if (UserAcct.Enabled == false)
                            {
                                    //Check if we are just logging deletions or actually deleting the profiles
                                    logger.Info("Deleting Profile: " + UserAccount);
                                    //Library.WriteErrorLog("Deleting Profile: " + UserAccount);
                                    //ProcessStartInfo startInfo = new ProcessStartInfo();
                                    
                                    //user32.dll not working.  Doesn't remove the folder or registry entries.
                                    //Library.DeleteProfile(newkey.GetValue("Sid").ToString(), newkey.GetValue("ProfileImagePath").ToString(), System.Environment.MachineName);
                                    //Process.Start(AppDomain.CurrentDomain.BaseDirectory + "DelProf2.exe", "/u /id2:" + UserAccount + " >>DelProf2_" + DateTime.Today.ToString("yyyyMMdd") + ".log");
                                    string procargs = "DelProf2.exe /u /id2:" + UserAccount + " >>\"" + Properties.Settings.Default.LogPath + "DelProf2_" + DateTime.Today.ToString("yyyyMMdd") + ".log\"";

                                    //startInfo.WorkingDirectory = Properties.Settings.Default.EXEPath;
                                    //startInfo.FileName = "DelProf2.exe";
                                    //startInfo.Arguments = procargs;
                                    //string EXEstring = startInfo.FileName + startInfo.Arguments;

                                    //logger.Trace("WorkingDirectory:" + startInfo.WorkingDirectory);
                                    //logger.Trace("Arguments:" + startInfo.Arguments);
                                    //logger.Trace("Filename:" + startInfo.FileName);
                                    logger.Trace("procargs:" + procargs);




                                    //Library.WriteErrorLog("WorkingDirectory:" + startInfo.WorkingDirectory);
                                    //Library.WriteErrorLog("Arguments:" + startInfo.Arguments);
                                    //Library.WriteErrorLog("Filename:" + startInfo.FileName);
                                    //logger.Trace(procargs);


                                    if (FirstUser == true)
                                    {
                                        logger.Trace("FirstUser:"+FirstUser.ToString());
                                        Library.WriteBatchFile("cd \"" + Properties.Settings.Default.EXEPath + "\"", true);
                                    }

                                    logger.Trace("procargs:" + procargs);
                                    Library.WriteBatchFile(procargs, false);

                                    

                                    FirstUser = false;
                                    logger.Trace("FirstUser:" + FirstUser.ToString());


                                    if (Properties.Settings.Default.DelProfile == "false")
                                    {
                                        logger.Trace("Properties.Settings.Default.DelProfile:" + Properties.Settings.Default.DelProfile.ToString());
                                        logger.Trace("Would have deleted this account:" + UserAccount);
                                    }

                            }
                            else
                                {
                                    //An account was found but it is enabled
                                    logger.Trace("The account was found but is enabled:" + UserAccount);
                                }
                        }
                        else
                        {
                                //The account was not found in the domains so advise support that this account may be able to be deleted manually
                                //since in the past ICON deleted old accounts from Active Directory instead of disabling them
                                //logger.Trace("The account was not found in the domains.  Check to see if this is a super old account.");
                        }

                    }
                }

            }  //end the loop of users

                //All users have been processed start the batch file if DelProfile is true
            if (Properties.Settings.Default.DelProfile == "true")
            {
                    logger.Trace("Properties.Settings.Default.DelProfile:" + Properties.Settings.Default.DelProfile.ToString());
                    logger.Trace("BATPath:" + Properties.Settings.Default.BATPath + "WindowsProfileCleaner.bat");
                    if (FirstUser == false)
                    {
                        //If FirstUser has been switched to false then we discovered at least one user that needs to be deleted
                        Process.Start(Properties.Settings.Default.BATPath + "WindowsProfileCleaner.bat");
                    }
                    else
                    {
                        //If FirstUser is still set to true then we didn't find any users who were disabled in AD
                        logger.Trace("Properties.Settings.Default.DelProfile:" + Properties.Settings.Default.DelProfile.ToString());
                        logger.Trace("BATPath:" + Properties.Settings.Default.BATPath + "WindowsProfileCleaner.bat");
                        logger.Trace("No users found that were disabled.");
                    }
                    
            }
            else
            {
                    logger.Info("Properties.Settings.Default.DelProfile" + Properties.Settings.Default.DelProfile.ToString());
            }

            } //end try
            catch (Exception ex)
            {
                //do nothing
                //need to add email or Application Log
                logger.Error(ex, "An error occurred:");
                //Library.WriteErrorLog("Error:" + ex.Message);
            }
        }

        }
    } //public partial class Scheduler : ServiceBase

