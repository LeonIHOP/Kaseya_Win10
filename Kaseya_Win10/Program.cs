using System.IO;
using System.Linq;
using System.Net;
using System;
using Microsoft.Win32;
using System.Collections.Generic;

namespace Kaseya_Win10
{



    class Program
    {

        #region Constants
        static string LogPath = "C:\\Source\\Kaseya\\KaseyaInstall.log"; //log file home folder
        static string cKey = "antimasker";
        static string Log;
        static string uName = "Administrator";
        static string store;
        

        //static string sPane;

        #endregion

        /// <summary>
        ///  Entry point into this application 
        ///  This app installs Kaseya by downloading the installer KcsSetup.exe and launching it
        ///  After a succesfull install, the app deletes KcsSetup.exe if it is in c:\startup directory
        /// </summary>
        static void Main() // Entry point into this app
        {
            
            try
            {
                Directory.CreateDirectory("C:\\Source\\Kaseya"); //creating directories for the installer executable Kaseya_Win10.exe
                Directory.CreateDirectory("D:\\Source\\Kaseya"); //creating directories for the installer executable Kaseya_Win10.exe

              
            }
             catch(Exception ex) {

                DE_Helpers.DE_FileManager.Log("***** Kaseya Install Failed ***** " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                DE_Helpers.DE_FileManager.Log("***** Failed to create directories for Kaseya_Win10.exe C:\\Source\\Kaseya D:\\Source\\Kaseya ***** " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
            }
            DE_Helpers.DE_FileManager.Log("***** Kaseya Install Script 1.4 *****", DE_Helpers.DE_FileManager.LogEntryType.Note);
            try
            {
                store = StoreNumber("IHOP"); // Pass in 'IHOP' string and recieve a store number from entry in the C:\B50\poll.bat file
                PollCheck(store); //call PollCheck and pass in the store number
            }
             catch(Exception ex) {

                DE_Helpers.DE_FileManager.Log("***** Kaseya Install Failed ***** " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                DE_Helpers.DE_FileManager.Log("***** Failed on PollCheck or AgentCheck *****", DE_Helpers.DE_FileManager.LogEntryType.Note);
            }

        }

        /// <summary>
        ///  Check for the presence of poll.bat
        ///  If exists, do a Franchise or Dine install of Kaseya
        ///  If mot then do clean up and exit app
        /// </summary>
        static void PollCheck(string store)
        {
            DE_Helpers.DE_FileManager.Log("Beginning poll.bat process", DE_Helpers.DE_FileManager.LogEntryType.Note);
            FrInstall(); // call FrInstall() to install the SERVER agent
        }

     

        /// <summary>
        ///  check if poll.bat is present and call Poll function
        /// </summary>
        static string StoreNumber(string concept)
        {
            if (File.Exists("C:\\B50\\poll.bat")) // if poll.bat exists
            {
                DE_Helpers.DE_FileManager.Log("looking for store number in poll.bat", DE_Helpers.DE_FileManager.LogEntryType.Note);
                string result = Poll(concept); // pass concept variable to Poll function
                DE_Helpers.DE_FileManager.Log("Found store " + result, DE_Helpers.DE_FileManager.LogEntryType.Note);
                return result; // return store number
            }
            else {
                DE_Helpers.DE_FileManager.Log("Poll.bat does not exists", DE_Helpers.DE_FileManager.LogEntryType.Note);
                return ""; // return blank if poll.bat is not found
            }
               

        }


        /// <summary>
        /// read poll.bat and extract store number or rturn a blank
        /// </summary>
        static string Poll(string concept)
        {
            string PollLine;
            string [] PollLineArray;
            if (concept == "IHOP")
            {
                try
                {
                    PollLine = DE_Helpers.DE_FileManager.ReadAllFromFile(@"C:\B50\poll.bat"); // read the one line of text in poll.bat
                    int positionOfNewLine = PollLine.IndexOf("\r\n");
                    if (positionOfNewLine >= 0)
                       PollLine = PollLine.Substring(0, PollLine.IndexOf("\r\n")).Trim();  // trim any leading and trailing blanks

                    PollLineArray = PollLine.Split(' ');
                    //var lastOperatorIndex = PollLine.LastIndexOf(" "); // position of the last blank character in this string?
                    //string storeNumber = PollLine.Substring(lastOperatorIndex,PollLine.Length - lastOperatorIndex).Trim(); // extract all characters  after the last space (store number) 
                    string storeNumber = PollLineArray[3]; // the store number should be in the 3rd position (0-1-2-3) 
                    return storeNumber;
                }
                 catch(Exception ex) {
                    DE_Helpers.DE_FileManager.Log("Error reading poll.bat " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                    return "";
                }

            }
            else return ""; //will never occur because 'IHOP' is hardcoded
            
        }

              
        /// <summary>
        /// Download store number specific Kaseya install executable 
        /// Launch this executable to install Kaseya
        /// If no match then invoke  DEInstall() to install Kaseya without a store number
        /// backup Kaseya_Win10.exe, delete it if its in c:\startup and exit this application
        /// </summary>
        static void FrInstall()
        {
                List<KaseyaAgent> KasayaAgentLinkedList = null;
                string CurrentDir = AppDomain.CurrentDomain.BaseDirectory; // place the path of the current directory into CurrentDir                     
                string csv_localPath = CurrentDir + @"DE_IHOP_CC_Agents.csv"; // string csv contains the full path of the DE_IHOP_CC_Agents.csv file
                try
                {
                    string newURL_CSV = ("ftp://ftp.ihop.com/Kaseya/DE_IHOP_CC_Agents.csv"); // FTP path to download KcsSetup.exe  
                    FtpWebRequest request = (FtpWebRequest)WebRequest.Create(newURL_CSV); // request is an object that is used for making FTP calls
                    request.Method = WebRequestMethods.Ftp.DownloadFile;
                    request.Credentials= new NetworkCredential("posftp", "befre!2O"); // FTP credentials
                    FtpWebResponse response = (FtpWebResponse)request.GetResponse(); // respomce is an object used for FTP response 
                    Stream responseStream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(responseStream);
                    string csv_content = reader.ReadToEnd();
                    DE_Helpers.DE_FileManager.WriteToFile(csv_content, csv_localPath); //write response content to file
                    DE_Helpers.DE_FileManager.Log("store list csv file Download complete", DE_Helpers.DE_FileManager.LogEntryType.Note);
                    reader.Close(); //close FTP
                    response.Close(); //close FTP
                }
                 catch(Exception ex)
                {
                    DE_Helpers.DE_FileManager.Log("store list csv file download failed, using local CSV file " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                }

                try
                {
                    string storeFromcompName = null;
                    var query = from currentLine in File.ReadAllLines(csv_localPath) // Linq query statement to read and parse the DE_IHOP_CC_Agents.csv file, x variable points to the current line
                            let column = currentLine.Split(',') // use ',' as the delimiter to separate the columns in the curent record
                            select new KaseyaAgent(column[0], column[1], column[2]); //column[0] = Store, column[1] = Agent, column[2] = Site

                    KasayaAgentLinkedList = query.Where(m => (m.Store.Length > 0) && (m.Site.Length > 0)).ToList(); // populate a linked list using the Linq variable 'query', discard missing data
                    storeFromcompName = Environment.MachineName.ToString().Trim();
                    
                    if (!((FindStoreInCSV(KasayaAgentLinkedList, store)) || (FindStoreInCSV(KasayaAgentLinkedList, storeFromcompName.Substring(storeFromcompName.Length - 4) )))) // use poll.bat store number or get store number from the computer name
                    {
                        DE_Helpers.DE_FileManager.Log("No match for StoreNUmber " + store + "in DE_IHOP_CC_Agents.csv ", DE_Helpers.DE_FileManager.LogEntryType.Note);
                        DEInstall(); // if no match is found between the storeId in poll.bat and a record in DE_IHOP_CC_Agents.csv do a DEinstall
                    }
                
                }
                catch(Exception ex)
                {
                    DE_Helpers.DE_FileManager.Log("Read and parse of DE_IHOP_CC_Agents.csv failed " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                    DEInstall(); // do a DEinstall due to a bad DE_IHOP_CC_Agents.csv file
            }
        }
           

        /// <summary>
        ///  DineEquity Installation
        /// Download DineEquity Kaseya installer and install Kaseya  
        /// </summary>
        static void DEInstall()
        {
            int retries = 0;
            int maxRetries = 3;
            bool downloaded = false;

            
            DE_Helpers.DE_FileManager.Log("Downloading generic DineEquity Kaseya Agent...", DE_Helpers.DE_FileManager.LogEntryType.Note);
           // WebClient wc = new WebClient(); // wc is an object that is used for making WWW calls

            while (!downloaded && (retries < maxRetries))
            {

                try
                {
                    using (WebClient wc = new WebClient())
                    {
                        DE_Helpers.DE_FileManager.Log("DineEquity KcsSetup.exe download is in progress. URL is "  + "https://cc.rosnet.com/mkDefault.asp?id=58224222" , DE_Helpers.DE_FileManager.LogEntryType.Note);
                        wc.DownloadFile("https://cc.rosnet.com/mkDefault.asp?id=58224222", @"c:\temp\KcsSetup.exe"); // download KcsSetup.exe which is a Kaseya installer from Rosnet to c:\temp directory
                        downloaded = true;
                        wc.Dispose();
                        DE_Helpers.DE_FileManager.Log("DineEquity KcsSetup.exe Download complete. Installing generic DineEquity Kaseya agent...", DE_Helpers.DE_FileManager.LogEntryType.Note);
                    }
                }
                catch (Exception ex)
                {
                    retries++;
                    DE_Helpers.DE_FileManager.Log("Error occurred downloading generic DineEquity KcsSetup.exe. Retry number " + retries.ToString() + " " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                    System.Threading.Thread.Sleep(100000);
                }

            }
                       
            if (downloaded == true)
            {
                DE_Helpers.DE_FileManager.Log("Starting the generic DineEquity Kaseya install...", DE_Helpers.DE_FileManager.LogEntryType.Note);
                try
                {
                    string commandToRun = @"c:\temp\KcsSetup.exe";
                    DE_Helpers.DE_FileManager.OperatingSystem os = DE_Helpers.DE_FileManager.GetOperatingSystemVersion();
                    string output = DE_Helpers.DE_FileManager.RunCommandAsAdmin(commandToRun, 600, os); // Run the Kaseya setup, wait at most 10 minutes 600 seconds for the install to complete
                    DE_Helpers.DE_FileManager.Log("Installed generic Dine Kaseya agent succesfully...", DE_Helpers.DE_FileManager.LogEntryType.Note);
                }
                catch (Exception ex)
                {
                    DE_Helpers.DE_FileManager.Log("Error occured running KcsSetup.exe  " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                    TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if it is in c:\startup and exit
                }


                if (File.Exists(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Kaseya\Kaseya Agent.lnk")) // path to the Kaseya shortcut file
                    DE_Helpers.DE_FileManager.Log("Succesfully confirmed Generic Kaseya agent install... ", DE_Helpers.DE_FileManager.LogEntryType.Note);
                else
                    DE_Helpers.DE_FileManager.Log("Failed to confirm Generic Kaseya agent install... ", DE_Helpers.DE_FileManager.LogEntryType.Note);
                TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if it is in c:\startup and exit
                                          

                DE_Helpers.DE_FileManager.Log("generic DineEquity Kaseya Install Complete.", DE_Helpers.DE_FileManager.LogEntryType.Note);
                
            }
            else
            {
                DE_Helpers.DE_FileManager.Log("Failed to download generic DineEquity KcsSetup.exe after retries " + retries.ToString(), DE_Helpers.DE_FileManager.LogEntryType.Note);
            }
            TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if it is in c:\startup and exit
        }

        /// <summary>
        /// Check Directory, backup, and Log
        /// Turn on Windowws Firewall
        /// Copy the Kaseya installer to c:\source\Kaseya
        /// Copy the installer to D:\ startup directory
        /// Delete the installer if it is running form c:\ startup directory
        /// </summary>
        static void TurnOnFireWallAndCopyLocal()
        {
            string commandToRun;
            string tempFileDir = AppDomain.CurrentDomain.BaseDirectory;
            try
            {
                commandToRun = @"netsh.exe /c Advfirewall set allprofiles state on"; // turn on windows firewall command
                DE_Helpers.DE_FileManager.OperatingSystem os = DE_Helpers.DE_FileManager.GetOperatingSystemVersion();
                string output = DE_Helpers.DE_FileManager.RunCommandAsAdmin(commandToRun, 600, os); // run the 'turn on firewall' command in the command line
                DE_Helpers.DE_FileManager.Log("Advfirewall set allprofiles state on ", DE_Helpers.DE_FileManager.LogEntryType.Note);
            }
             catch(Exception ex)
            {
                DE_Helpers.DE_FileManager.Log("Error executing command: Advfirewall set allprofiles state on" + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                
            }

            try
            {
                DE_Helpers.DE_FileManager.Log("Copying to C:\\Source\\Kaseya and secondary...", DE_Helpers.DE_FileManager.LogEntryType.Note);
                DE_Helpers.DE_FileManager.CopyAndCreateAsAdmin(AppDomain.CurrentDomain.BaseDirectory + @"\Kaseya_Win10.exe", @"C:\Source\Kaseya\Kaseya_Win10.exe"); // backup to  C:\Source\Kaseya
                DE_Helpers.DE_FileManager.CopyAndCreateAsAdmin(@"C:\Source\Kaseya\Kaseya_Win10.exe", @"D:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Kaseya_Win10.exe"); //copying to d: startup so in failure recoevry to D: we will run Kaseya on startup
            }
            catch(Exception ex)
            {
                DE_Helpers.DE_FileManager.Log("Could not Copy to C:\\Source\\Kaseya and D:\\startup... " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
            }
            // SelfDelete
            if (AppDomain.CurrentDomain.BaseDirectory == "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\") // if this executable is in the startup directory
            {
                DE_Helpers.DE_FileManager.Log("Install ran from the startup directory, self deleting and exiting...", DE_Helpers.DE_FileManager.LogEntryType.Note);
                // create delete_Kaseya batch file that will delete the Kaseya install .exe because we only want the installer to run once
                //1. wait for 10 seconds before performing the delete
                //2. change directory to the startup directory
                //3. delete DE_IHOP_CC_Agents.csv
                //4. delete Kaseya_Win10.exe
                //5. delete DE_Helpers.dll
                Directory.CreateDirectory("C:\\Temp"); //crete c:\temp if needed
                try
                {
                    string TextToWrite = "timeout /T 10 /NOBREAK > NUL " + Environment.NewLine + "cd /d " + "\"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\"" 
                        + Environment.NewLine + "del " + "\"Kaseya_Win10.exe\"";
                    if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + @"DE_IHOP_CC_Agents.csv"))
                        TextToWrite = TextToWrite + Environment.NewLine + "del " + "\"DE_IHOP_CC_Agents.csv\"";
                    if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + @"DE_Helpers.dll"))
                        TextToWrite = TextToWrite + Environment.NewLine + "del " + "\"DE_Helpers.dll\"";
                    try
                    {
                        DE_Helpers.DE_FileManager.Log("Creating C:\\Temp\\delete_Kaseya.bat ", DE_Helpers.DE_FileManager.LogEntryType.Note);
                        DE_Helpers.DE_FileManager.WriteToFile(TextToWrite, "C:\\Temp\\delete_Kaseya.bat"); // write the batch file commands to delete_Kaseya.bat

                    }
                    catch (Exception ex)
                    {
                        DE_Helpers.DE_FileManager.Log("unable to create C:\\Temp\\delete_Kaseya.bat " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                        AppExit(); //log and exit 
                    }

                    //DE_Helpers.DE_FileManager.OperatingSystem os = DE_Helpers.DE_FileManager.GetOperatingSystemVersion(); // assign OS version to os variable
                    //string output = DE_Helpers.DE_FileManager.RunCommandAsAdmin(commandToRun,1,os);
                    DE_Helpers.DE_FileManager.Log("invoking C:\\Temp\\delete_Kaseya.bat ... ", DE_Helpers.DE_FileManager.LogEntryType.Note);
                    try
                    {
                        commandToRun = "C:\\Temp\\delete_Kaseya.bat"; // create the command that will execute this batch file
                        System.Diagnostics.Process.Start(commandToRun); // run the batch file command at the command line
                    }
                    catch (Exception ex) {
                        DE_Helpers.DE_FileManager.Log("unable to run C:\\Temp\\delete_Kaseya.bat " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                        AppExit(); //log and exit
                    }
                }
                catch (Exception ex)
                {
                    DE_Helpers.DE_FileManager.Log("unable to delete from startup... " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                }
                System.Environment.Exit(0); // shut down this app
            }

            AppExit(); //log and exit 
        }

        private static bool FindStoreInCSV(List<KaseyaAgent> KasayaAgentLinkedList, string store)
        {
            bool storeFound = false;
            DE_Helpers.DE_FileManager.Log("Searching for Store number " + store, DE_Helpers.DE_FileManager.LogEntryType.Note);
            KasayaAgentLinkedList.ForEach(m => // visit each node in the KasayaAgent linked list, the current node is pointed to by m variable 
            {

                if (m.Store == store) // if a match is found between the storeId in poll.bat and a record in DE_IHOP_CC_Agents.csv do a Franchise install
                {
                    string Agent = m.Agent;
                    string site = m.Site;
                    string Text = "install";
                    string[] FoundWord; //string array that will contain HTML code substrings
                    int retries = 0;
                    int maxRetries = 3; // retry 3 times at most
                    bool downloaded = false;
                    storeFound = true;
                    try
                    {

                        try
                        {
                            string source = new System.Net.WebClient().DownloadString(site); //pull down the HTML source from the download page so we can extract the direct download URL from it
                            int KeyWord = source.IndexOf(Text, 1); // find the install string after href on the ROSnet download page for this restaurant

                            if (KeyWord != 0) // no error
                            {

                                FoundWord = source.Substring(KeyWord).Split('"').ToArray(); // split the HTML source code into sub-strings and insert into the FoundWord array
                                string newURL = ("https://cc.rosnet.com/" + FoundWord[0]); // construct the URL to download KcsSetup.exe by appending KcsSetup.exe path to Rosnet URL 
                                DE_Helpers.DE_FileManager.Log(" Downloading Franchise KcsSetup.exe ...", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                while (!downloaded && (retries < maxRetries))
                                {

                                    try
                                    {
                                        using (WebClient wc = new WebClient())
                                        {
                                            DE_Helpers.DE_FileManager.Log("Franchise KcsSetup.exe download is in progress. URL is " + newURL, DE_Helpers.DE_FileManager.LogEntryType.Note);
                                            wc.DownloadFile(newURL, @"c:\temp\KcsSetup.exe"); // download KcsSetup.exe which is a Kaseya installer from Rosnet to c:\temp directory
                                            downloaded = true;
                                            wc.Dispose();
                                            DE_Helpers.DE_FileManager.Log("Franchise KcsSetup.exe Download complete. Installing Franchise Kaseya agent...", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        retries++;
                                        DE_Helpers.DE_FileManager.Log("Error occurred downloading Franchise KcsSetup.exe. Retry number " + retries.ToString() + " " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                                        System.Threading.Thread.Sleep(100000);
                                    }

                                }
                                if (downloaded == true)
                                {
                                    DE_Helpers.DE_FileManager.Log("Starting the Kaseya install...", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                    try
                                    {
                                        string commandToRun = @"c:\temp\KcsSetup.exe";
                                        DE_Helpers.DE_FileManager.OperatingSystem os = DE_Helpers.DE_FileManager.GetOperatingSystemVersion(); // assign OS version to os variable
                                        string output = DE_Helpers.DE_FileManager.RunCommandAsAdmin(commandToRun, 600, os); // Run the Kaseya setup, wait at most 10 minutes 600 seconds for the install to complete
                                                                                                                            //string output = DE_Helpers.DE_FileManager.RunCommand(commandToRun, 600); // Run the Kaseya setup, wait at most 10 minutes 600 seconds for the install to complete
                                        DE_Helpers.DE_FileManager.Log("Installed Franchise Kaseya agent ..." + output, DE_Helpers.DE_FileManager.LogEntryType.Note);
                                    }
                                    catch (Exception ex)
                                    {
                                        DE_Helpers.DE_FileManager.Log("Install exited. No success. Error occured running KcsSetup.exe... " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                                        TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if its in c:\startup and exit this application
                                    }

                                    if (File.Exists(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Kaseya\Kaseya Agent.lnk")) // path to the Kaseya shortcut file
                                        DE_Helpers.DE_FileManager.Log("Succesfully confirmed Franchise Kaseya agent install... ", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                    else
                                        DE_Helpers.DE_FileManager.Log("Failed to  confirm Franchise Kaseya agent install... ", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                    TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if it is in c:\startup and exit
                                }
                                else // could not download KcsSetup.exe
                                {
                                    DE_Helpers.DE_FileManager.Log("Failed to download Franchise KcsSetup.exe after retries " + retries.ToString(), DE_Helpers.DE_FileManager.LogEntryType.Note);
                                }
                            }
                            else
                            {
                                DE_Helpers.DE_FileManager.Log("Error occured, Could not locate the install string after href on the ROSnet download page for this restaurant ", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                
                            }
                        }
                        catch (Exception ex)
                        {
                            DE_Helpers.DE_FileManager.Log("Error occured downloading HTML source from the download page " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                            
                        }
                    }
                    catch (Exception ex)
                    {
                        DE_Helpers.DE_FileManager.Log("Error occured during install " + ex.Message, DE_Helpers.DE_FileManager.LogEntryType.Note);
                        
                    }
                }

            });
            return storeFound;
        }
        /// <summary>
        /// Log status and exit application
        /// </summary>
        static void AppExit()
        {

            DE_Helpers.DE_FileManager.Log("Exiting install.", DE_Helpers.DE_FileManager.LogEntryType.Note);
            DE_Helpers.DE_FileManager.Log("************************************", DE_Helpers.DE_FileManager.LogEntryType.Note);
            System.Environment.Exit(0); //shut down this app

        }

        public class KaseyaAgent // CSV data structure
        {
            public string Store { get; set; }
            public string Agent { get; set; }
            public string Site { get; set; }

            public KaseyaAgent(string Store, string Agent, string Site)
            {
                this.Store = Store;
                this.Agent = Agent;
                this.Site = Site;
            }
        }
    }

}

