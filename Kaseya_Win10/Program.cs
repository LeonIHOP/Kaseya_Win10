using System.IO;
using System.Linq;
using System.Text;
using System.Net;
using System.Diagnostics;
using System.Security;

namespace Kaseya_Win10
{



    class Program
    {

        #region Constants
        static string LogPath = "C:\\Source\\Kaseya\\KaseyaInstall.log";
        static string cKey = "antimasker";
        static string Log;
        static string uName = "Administrator";
        static string store;
        static string sPane;
        #endregion


        static void Main()
        {
           Directory.CreateDirectory("C:\\Source\\Kaseya");
           //Directory.CreateDirectory("D:\\Source\\Kaseya");

           store = StoreNumber("IHOP");
           sPane = sMask(cKey);
            
            PollCheck(store);
            AgentCheck();

        }

        
        static void PollCheck(string store)
        {
           switch (store)
           {
                case "":
                    //        _FileWriteLog($LogPath, "Error: poll.bat not found or could not be read.");
                    //CheckDir();
                    break;
                case "XXXX":
                    FrInstall(); // pointed to FrInstall() to install the SERVER agent, and becuase _StoreNumber only plays nice with numbers
                    break;
                    //     Case @error:
                    //$er = @error;
                    //     _FileWriteLog($LogPath, "Error occurred reading the poll.bat: " & $er & " - " & @extended & " - " & _WinAPI_GetLastErrorMessage());
                    //     CheckDir();
                    error:;
                    //CheckDir();
            }

        }

        static void AgentCheck()
        {

            //RegistryKey rk = baseRegistryKey;
            //RegistryKey sk1 = rk.OpenSubKey(@"\WOW6432NODE\KASEYA\AGENT");

            //var iQuery = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432NODE\KASEYA\AGENT",, null);

            //   If ($iQuery == "")
            FrInstall();
            //   Else
            //   { 
            //       $iAgent = RegRead("HKLM64\\SOFTWARE\\WOW6432NODE\\KASEYA\\AGENT\" & $iQuery, "MachineID");
            //       _FileWriteLog($LogPath, "Kaseya agent " & $iAgent & " already installed.");
            //       CheckDir();
            //   }


        }

        static string StoreNumber(string concept)
        {
            if (File.Exists("C:\\B50\\poll.bat"))
            {
                string result = Poll(concept);
                return result;
            }
            else
            return "";
                  
        }

        static string Poll(string concept)
        {
            //if (concept == "IHOP")
            //{

            //    number = @ComSpec & " /C" & 'for /f "tokens=4 delims== " %i in (c:\b50\Poll.bat) do (echo %i)';
            //}

            //      If StringLeft($number, 1) = " " Then $number = " " & $number

            //      $nPid = Run(@ComSpec & " /c" & $number, "", @SW_HIDE, 8), $sRet = ""

            //      If @error Then Return @error

            //       ProcessWait($nPid)

            //      While 1
            //$sRet &= StdoutRead($nPid)

            //      If @error Or(Not ProcessExists($nPid)) Then ExitLoop

            //      Return StringTrimRight(StringRight($sRet, 6), 2); trims two spaces off the right to avoid catching the " )" in the echo
            return "9976";
        }

        static string sMask(string cText)
        {
            int i;
            char[] aArray;
            byte [] key = {0x2f, 0x4d, 0x7, 0x21, 0x5, 0x45, 0x4a, 0x1d, 0xa, 0x42 };

            byte[] sb = Encoding.ASCII.GetBytes(cText).ToArray();

            aArray =  new char[cText.Length];
            for (i = 0; i < (cText.Length - 1); i++)
            {
                aArray[i] = (char)(key[i] ^ sb[i]);
            }

            string  built = new string(aArray);
            built = built.Replace("\0", "0");
            return built;


        }
        static void FrInstall()
        {
            string tempFileDir = Path.GetTempPath();
            string csv = tempFileDir + @"DE_IHOP_CC_Agents.csv";


            var query = from x in File.ReadAllLines(csv)
                                        let p = x.Split(',')
                                        select new KasayaAgent(p[0], p[1], p[2]); //p[0] = Store, p[1] = Agent, p[2] = Site

            var KasayaAgenrt = query.ToList();

            KasayaAgenrt.ForEach(m =>
            {
                if (m.Store == store)
                {
                    string iAgent = m.Agent;
                    string site = m.Site;
                    string Text = "install";
                    string[] FoundWord;
                    string source = new System.Net.WebClient().DownloadString(site);
                    int KeyWord = source.IndexOf(Text, 1);
                    if (KeyWord != 0)
                    {

                        FoundWord = source.Substring(KeyWord).Split('"').ToArray();
                        string newURL = ("https://cc.rosnet.com/" + FoundWord[0]);
                        WebClient wc_ = new WebClient();
                        wc_.DownloadFile(newURL, @"c:\temp\KcsSetup.exe");
                        //error checking and log file update here
                        Process processKaseyaInstall = new Process();
                        processKaseyaInstall.StartInfo.FileName = @"c:\temp\KcsSetup.exe";
                        processKaseyaInstall.StartInfo.UserName = uName;
                        processKaseyaInstall.StartInfo.Domain = System.Environment.MachineName;
                        var securePane = new SecureString();
                        foreach (char c in sPane)
                        {
                            securePane.AppendChar(c);
                        }
                        processKaseyaInstall.StartInfo.Password = securePane;
                        processKaseyaInstall.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        processKaseyaInstall.Start();
                        processKaseyaInstall.WaitForExit();
                        CheckDir();
                    }
                }

            });


            }
    }

    public class KasayaAgent // CSV data structure
    {
        public string Store { get; set; }
        public string Agent { get; set; }
        public string Site { get; set; }

        public KasayaAgent(string Store, string Agent, string Site)
        {
            this.Store = Store;
            this.Agent = Agent;
            this.Site = Site;
        }
    }


}
