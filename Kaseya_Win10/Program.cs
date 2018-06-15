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

                store = StoreNumber("IHOP"); // Pass in 'IHOP' string and recieve a store number from entry in the C:\B50\poll.bat file

                DE_Helpers.DE_FileManager.Log("***** Kaseya Install Script 1.4 *****", DE_Helpers.DE_FileManager.LogEntryType.Note);
                PollCheck(store); //call PollCheck and pass in the store number
                AgentCheck(); // Call AgentCheck to install Kaseya if it is not installed
            }
            catch {

                DE_Helpers.DE_FileManager.Log("***** Kaseya Install Failed *****", DE_Helpers.DE_FileManager.LogEntryType.Note);
                DE_Helpers.DE_FileManager.Log("***** Failed to create directories for Kaseya_Win10.exe *****", DE_Helpers.DE_FileManager.LogEntryType.Note);
            }

           

        }

        /// <summary>
        ///  Check for the presence of poll.bat
        ///  If exists, do a Franchise install of Kaseya
        ///  If mot then do clean up and exit app
        /// </summary>
        static void PollCheck(string store)
        {
            DE_Helpers.DE_FileManager.Log("Beginning poll.bat process", DE_Helpers.DE_FileManager.LogEntryType.Note);
            switch (store)
            {
                case "": // store number passed in is blank
                    DE_Helpers.DE_FileManager.Log("Error: poll.bat not found or could not be read.", DE_Helpers.DE_FileManager.LogEntryType.Note);
                    TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if it is in c:\startup and exit this program
                    break;
                case "XXXX": // store number passed in is XXXX
                    FrInstall(); // call FrInstall() to install the SERVER agent
                    break;

                default: //error reading poll.bat
                    DE_Helpers.DE_FileManager.Log("Error occurred reading the poll.bat", DE_Helpers.DE_FileManager.LogEntryType.Note);
                    TurnOnFireWallAndCopyLocal(); // // backup Kaseya_Win10.exe, delete it if it is in c:\startup and exit this program
                    break;
            }

        }

        /// <summary>
        ///  Perform a Franchise install if Kaseya is not installed, else back up and exit application
        /// </summary>
        static void AgentCheck() // agent check and install
        {

            using (RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\KASEYA\\AGENT")) // read this registry entry
            {
                DE_Helpers.DE_FileManager.Log("Checking to see if Kaseya is already installed.", DE_Helpers.DE_FileManager.LogEntryType.Note);
                if (key != null) // if it is populated
                {
                    var iQuery = key.GetValue("DriverControl"); // get value of the 'DriverControl' field
                    if (iQuery.ToString() == "") // if it is empty
                    {
                        FrInstall(); // do a Franchise install
                    }
                    else { // Kaseya is already installed

                        var iAgent = key.GetValue("DriverControl\\"+ iQuery+ "MachineID");
                        DE_Helpers.DE_FileManager.Log("Kaseya agent " + iAgent + " already installed.", DE_Helpers.DE_FileManager.LogEntryType.Note);
                        TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if it is in c:\startup and exit this program
                    }

                }
            }
           
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
                return result; // return store number
            }
            else
                return ""; // return blank if poll.bat is not found

        }


        /// <summary>
        /// read poll.bat and extract store number or rturn a blank
        /// </summary>
        static string Poll(string concept)
        {
            string PollLine;
            if (concept == "IHOP")
            {
                try
                {
                    PollLine = DE_Helpers.DE_FileManager.ReadAllFromFile(@"C:\B50\poll.bat"); // read the one line of text in poll.bat
                    PollLine.Trim(); // trim any leading and trailing blanks
                    string storeNumber = PollLine.Substring(PollLine.Length - 4); // extract the last 4 characters, they are the store number
                    DE_Helpers.DE_FileManager.Log("Found store" + storeNumber, DE_Helpers.DE_FileManager.LogEntryType.Note);
                    return storeNumber;
                }
                catch {
                    DE_Helpers.DE_FileManager.Log("Error reading poll.bat", DE_Helpers.DE_FileManager.LogEntryType.Note);
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
                    DE_Helpers.DE_FileManager.Log("FTP Download complete", DE_Helpers.DE_FileManager.LogEntryType.Note);
                    reader.Close(); //close FTP
                    response.Close(); //close FTP
                }
                catch
                {
                    DE_Helpers.DE_FileManager.Log("FTP Download failed, using local CSV file", DE_Helpers.DE_FileManager.LogEntryType.Note);
                }

                try
                {
                    var query = from currentLine in File.ReadAllLines(csv_localPath) // Linq query statement to read and parse the DE_IHOP_CC_Agents.csv file, x variable points to the current line
                            let column = currentLine.Split(',') // use ',' as the delimiter to separate the columns in the curent record
                            select new KaseyaAgent(column[0], column[1], column[2]); //column[0] = Store, column[1] = Agent, column[2] = Site

                    KasayaAgentLinkedList = query.Where(m => (m.Store.Length > 0) && (m.Site.Length > 0) && (m.Agent.Length > 0)).ToList(); // populate a linked list using the Linq variable 'query', discard missing data
                }
                catch {

                    DE_Helpers.DE_FileManager.Log("read and parse of DE_IHOP_CC_Agents.csv failed", DE_Helpers.DE_FileManager.LogEntryType.Note);
                    return;

                }

                KasayaAgentLinkedList.ForEach(m => // visit each node in the KasayaAgent linked list, the current node is pointed to by m variable 
                {
                    
                        if (m.Store == store)
                        {
                            string Agent = m.Agent;
                            string site = m.Site;
                            string Text = "install";
                            string[] FoundWord; //string array that will contain HTML code substrings

                            try
                            {
                                string source = new System.Net.WebClient().DownloadString(site); //pull down the HTML source from the download page so we can extract the direct download URL from it
                                int KeyWord = source.IndexOf(Text, 1); // find the install string after href on the ROSnet download page for this restaurant

                                if (KeyWord != 0) // no error
                                {

                                    FoundWord = source.Substring(KeyWord).Split('"').ToArray(); // split the HTML source code into sub-strings and insert into the FoundWord array
                                    string newURL = ("https://cc.rosnet.com/" + FoundWord[0]); // construct the URL to download KcsSetup.exe by appending KcsSetup.exe path to Rosnet URL 
                                    WebClient wc = new WebClient(); // wc is an object that is used for making WWW calls
                                    try
                                    {
                                        wc.DownloadFile(newURL, @"c:\temp\KcsSetup.exe"); // download KcsSetup.exe which is a Kaseya installer from Rosnet to c:\temp directory
                                        DE_Helpers.DE_FileManager.Log("Download complete", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                        wc.Dispose();

                                    }
                                    catch
                                    {
                                        DE_Helpers.DE_FileManager.Log("Error occurred accessing URL" + newURL, DE_Helpers.DE_FileManager.LogEntryType.Note);
                                        TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if it is in c:\startup and exit
                                    }
                                    DE_Helpers.DE_FileManager.Log("Starting the Kaseya install...", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                    try
                                    {
                                        string commandToRun = @"c:\temp\KcsSetup.exe";
                                        DE_Helpers.DE_FileManager.OperatingSystem os = DE_Helpers.DE_FileManager.GetOperatingSystemVersion(); // assign OS version to os variable
                                        string output = DE_Helpers.DE_FileManager.RunCommandAsAdmin(commandToRun, 120, os); // Run the Kaseya setup, wait at most 2 minutes 600 seconds for the install to complete
                                        DE_Helpers.DE_FileManager.Log("Installing Kaseya agent...", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                    }
                                    catch
                                    {
                                        DE_Helpers.DE_FileManager.Log("Error occured during install", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                        TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if its in c:\startup and exit this application
                                    }

                                    using (RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\KASEYA\\AGENT"))
                                    {
                                        var iQuery = key.GetValue("DriverControl");
                                        var iAgent = key.GetValue("DriverControl\\" + iQuery + "MachineID");
                                        DE_Helpers.DE_FileManager.Log("Kaseya agent " + iAgent + " succesfully installed.", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                        DE_Helpers.DE_FileManager.Log("Install complete", DE_Helpers.DE_FileManager.LogEntryType.Note);
                                        TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if it is in c:\startup and exit
                                    }

                                }
                                else
                                {
                                    DE_Helpers.DE_FileManager.Log("Error occured during install", DE_Helpers.DE_FileManager.LogEntryType.Note);

                                }
                            }
                            catch
                            {
                                DE_Helpers.DE_FileManager.Log("Error occured during install", DE_Helpers.DE_FileManager.LogEntryType.Note);

                            }


                        }
                   
                });
            }
            

           

        /// <summary>
        ///  DineEquity Installation
        /// Download DineEquity Kaseya installer and install Kaseya  
        /// </summary>
        static void DEInstall()
        {
            DE_Helpers.DE_FileManager.Log("Downloading DineEquity Kaseya Agent...", DE_Helpers.DE_FileManager.LogEntryType.Note);
            WebClient wc = new WebClient(); // wc is an object that is used for making WWW calls
            try
            {
                wc.DownloadFile("https://cc.rosnet.com/mkDefault.asp?id=58224222", @"c:\temp\KcsSetup.exe"); // construct the URL to download KcsSetup.exe by appending KcsSetup.exe path to Rosnet URL 
                DE_Helpers.DE_FileManager.Log("Installing Kaseya agent...", DE_Helpers.DE_FileManager.LogEntryType.Note);
            }
            catch 
            {
                DE_Helpers.DE_FileManager.Log("Error occured during install", DE_Helpers.DE_FileManager.LogEntryType.Note);
                TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if it is in c:\startup and exit

            }

            DE_Helpers.DE_FileManager.Log("Starting the Kaseya install...", DE_Helpers.DE_FileManager.LogEntryType.Note);
            try
            {
                string commandToRun = @"c:\temp\KcsSetup.exe";
                DE_Helpers.DE_FileManager.OperatingSystem os = DE_Helpers.DE_FileManager.GetOperatingSystemVersion();
                string output = DE_Helpers.DE_FileManager.RunCommandAsAdmin(commandToRun, 120, os); // Run the Kaseya setup, wait at most 2 minutes 120 seconds for the install to complete
                DE_Helpers.DE_FileManager.Log("Installing Kaseya agent...", DE_Helpers.DE_FileManager.LogEntryType.Note);
            }
            catch
            {
                DE_Helpers.DE_FileManager.Log("Error occured during install", DE_Helpers.DE_FileManager.LogEntryType.Note);
                TurnOnFireWallAndCopyLocal(); // backup Kaseya_Win10.exe, delete it if it is in c:\startup and exit
            }
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\KASEYA\\AGENT")) // get the Kaseya agent value to use in the log 
            {
                var iQuery = key.GetValue("DriverControl");
                var iAgent = key.GetValue("DriverControl\\" + iQuery + "MachineID"); // read the Machine ID registry entry
                DE_Helpers.DE_FileManager.Log("Kaseya agent " + iAgent + " succesfully installed.", DE_Helpers.DE_FileManager.LogEntryType.Note);
            }

            DE_Helpers.DE_FileManager.Log("Install Complete.", DE_Helpers.DE_FileManager.LogEntryType.Note);
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
            catch
            {
                DE_Helpers.DE_FileManager.Log("Error executing command: Advfirewall set allprofiles state on", DE_Helpers.DE_FileManager.LogEntryType.Note);
                
            }

            DE_Helpers.DE_FileManager.Log("Copying to C:\\Source\\Kaseya and secondary...", DE_Helpers.DE_FileManager.LogEntryType.Note);
            DE_Helpers.DE_FileManager.CopyAndCreateAsAdmin(AppDomain.CurrentDomain.BaseDirectory + @"\Kaseya_Win10.exe", @"C:\Source\Kaseya\Kaseya_Win10.exe"); // backup to  C:\Source\Kaseya
            DE_Helpers.DE_FileManager.CopyAndCreateAsAdmin(@"C:\Source\Kaseya\Kaseya_Win10.exe", @"D:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Kaseya_Win10.exe"); //copying to d: startup so in failure recoevry to D: we will run Kaseya on startup
           
            // SelfDelete
            if (AppDomain.CurrentDomain.BaseDirectory == @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup") // if this executable is in the startup directory
            {
                DE_Helpers.DE_FileManager.Log("Install ran from the startup directory, self deleting and exiting...", DE_Helpers.DE_FileManager.LogEntryType.Note);
                // create delete_Kaseya batch file that will delete the Kaseya install .exe because we only want the installer to run once
                //1. wait for 10 seconds before performing the delete
                //2. change directory to the startup directory
                //3. delete Kaseya_Win10.exe
                string TextToWrite = "timeout /T 10 /NOBREAK > NUL " + Environment.NewLine + "cd /d " + "\"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\"" + Environment.NewLine + "del " + "\"Kaseya_Win10.exe\"";
                DE_Helpers.DE_FileManager.WriteToFile(TextToWrite , AppDomain.CurrentDomain.BaseDirectory + @"delete_Kaseya.bat"); // write the batch file commands to delete_Kaseya.bat
                commandToRun = AppDomain.CurrentDomain.BaseDirectory +  @"delete_Kaseya.bat"; // create the command that will execute this batch file
                System.Diagnostics.Process.Start(commandToRun); // run the batch file command at the command line
                System.Environment.Exit(1); // shut down this app
            }

            AppExit(); //log and exit 
        }

        /// <summary>
        /// Log status and exit application
        /// </summary>
        static void AppExit()
        {

            DE_Helpers.DE_FileManager.Log("Exiting install.", DE_Helpers.DE_FileManager.LogEntryType.Note);
            DE_Helpers.DE_FileManager.Log("************************************", DE_Helpers.DE_FileManager.LogEntryType.Note);
            System.Environment.Exit(1); //shut down this app

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

