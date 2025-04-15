using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.Win32;
using ActiveDs;


namespace ComputerVerifier
{
    class CVMain
    {
        struct ComputerLoginParams
        {
            public string strDisabledComputersLocation;
            public string strDisableUnusedComputers;
            public int intDaysSinceLastLogon;
            public int intDaysSinceDateCreated;
            public List<string> lstExcludePrefix;
            public List<string> lstExclude;
            public List<string> lstExcludeSuffix;
            public string strFilter;
        }

        struct CMDArguments
        {
            public string strQueryFilter;
            public bool bParseCmdArguments;
        }

        static bool funcContactServer(string strServerName)
        {
            bool bPingSuccess = false;
            bool bWMISuccess = false;

            Console.WriteLine("Contact start for {0}: {1}", strServerName, DateTime.Now.ToLocalTime().ToString("MMddyyy HH:mm:ss"));

            string strServerNameforWMI = "";
            // [Comment] Ping the server
            // [DebugLine] Console.WriteLine(); // Helper line just to make output clearer
            Console.WriteLine("Ping attempt for: " + strServerName);

            try
            {
                System.Net.NetworkInformation.Ping objPing1 = new System.Net.NetworkInformation.Ping();
                System.Net.NetworkInformation.PingReply objPingReply1 = objPing1.Send(strServerName);
                if (objPingReply1.Status.ToString() != "TimedOut")
                {
                    Console.WriteLine("Ping Reply: " + objPingReply1.Address + "     RTT: " + objPingReply1.RoundtripTime);
                    bPingSuccess = true;
                }
                else
                {
                    Console.WriteLine("Ping Reply: " + objPingReply1.Status);
                    bPingSuccess = false;
                }
            }
            catch (SystemException ex)
            {
                // [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());
                string strPingError = "An exception occurred during a Ping request.";
                // [DebugLine] System.Console.WriteLine(ex.Message);
                if (ex.Message == strPingError)
                {
                    Console.WriteLine("Ping Error. No ip address was found during name resolution.");
                }

                bPingSuccess = false;
            }

            // [Comment] Connect to the server via WMI
            // [DebugLine] Console.WriteLine(); // Helper line just to make output clearer
            Console.WriteLine("WMI connection attempt for: " + strServerName);

            System.Management.ConnectionOptions objConnOptions = new System.Management.ConnectionOptions();
            strServerNameforWMI = "\\\\" + strServerName + "\\root\\cimv2";

            // [DebugLine] Console.WriteLine("Construct WMI scope...");
            System.Management.ManagementScope objManagementScope = new System.Management.ManagementScope(strServerNameforWMI, objConnOptions);
            // [DebugLine] Console.WriteLine("Construct WMI query...");
            System.Management.ObjectQuery objQuery = new System.Management.ObjectQuery("select * from Win32_ComputerSystem");
            // [DebugLine] Console.WriteLine("Construct WMI object searcher...");
            System.Management.ManagementObjectSearcher objSearcher = new System.Management.ManagementObjectSearcher(objManagementScope, objQuery);
            Console.WriteLine("Get WMI data...");

            try
            {
                System.Management.ManagementObjectCollection objObjCollection = objSearcher.Get();

                foreach (System.Management.ManagementObject objMgmtObject in objObjCollection)
                {
                    Console.WriteLine("Hostname: " + objMgmtObject["Caption"].ToString());
                    bWMISuccess = true;
                }
            }
            catch (SystemException ex)
            {
                // [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());
                string strRPCUnavailable = "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)";
                // [DebugLine] System.Console.WriteLine(ex.Message);
                if (ex.Message == strRPCUnavailable)
                {
                    Console.WriteLine("WMI: Server unavailable");
                }
                bWMISuccess = false;
            }

            Console.WriteLine("Contact stop for {0}: {1}", strServerName, DateTime.Now.ToLocalTime().ToString("MMddyyy HH:mm:ss"));

            if (bPingSuccess & bWMISuccess)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        static void funcLogToEventLog(string strAppName, string strEventMsg, int intEventType)
        {
            string sLog;

            sLog = "Application";

            if (!EventLog.SourceExists(strAppName))
                EventLog.CreateEventSource(strAppName, sLog);

            //EventLog.WriteEntry(strAppName, strEventMsg);
            EventLog.WriteEntry(strAppName, strEventMsg, EventLogEntryType.Information, intEventType);

        } // LogToEventLog

        static void funcPrintParameterWarning()
        {
            Console.WriteLine("A parameter must be specified to run ComputerVerifier.");
            Console.WriteLine("Run ComputerVerifier -? to get the parameter syntax.");
        }

        static void funcPrintParameterSyntax()
        {
            Console.WriteLine("ComputerVerifier");
            Console.WriteLine();
            Console.WriteLine("Parameter syntax:");
            Console.WriteLine();
            Console.WriteLine("Use the following for the first parameter:");
            Console.WriteLine("-run                  required parameter");
            Console.WriteLine();
            Console.WriteLine("Use one of the following for the second parameter:");
            Console.WriteLine("-all                  for All Windows computer objects");
            Console.WriteLine("-allservers           for All Windows Server computer objects");
            Console.WriteLine("-allworkstations      for All Windows workstation computer objects");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("ComputerVerifier -run -all");
            Console.WriteLine("ComputerVerifier -run -allservers");
            Console.WriteLine("ComputerVerifier -run -allworkstations");
        } // funcPrintParameterSyntax

        static CMDArguments funcParseCmdArguments(string[] cmdargs)
        {
            CMDArguments objCMDArguments = new CMDArguments();

            try
            {
                objCMDArguments.bParseCmdArguments = false;

                if (cmdargs[0] == "-run" & cmdargs.Length > 1)
                {

                    for (int i = 1; i < cmdargs.Length; i++)
                    {
                        if (i == 1)
                        {
                            if (cmdargs[i] == "-all")
                            {
                                objCMDArguments.strQueryFilter = "-all";
                                objCMDArguments.bParseCmdArguments = true;
                            }

                            if (cmdargs[i] == "-allservers")
                            {
                                objCMDArguments.strQueryFilter = "-allservers";
                                objCMDArguments.bParseCmdArguments = true;
                            }

                            if (cmdargs[i] == "-allworkstations")
                            {
                                objCMDArguments.strQueryFilter = "-allworkstations";
                                objCMDArguments.bParseCmdArguments = true;
                            }
                        }
                    }
                }
                else
                {
                    objCMDArguments.bParseCmdArguments = false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                objCMDArguments.bParseCmdArguments = false;
            }

            return objCMDArguments;
        }

        static ComputerLoginParams funcParseConfigFile(CMDArguments objCMDArguments2)
        {
            ComputerLoginParams newParams = new ComputerLoginParams();

            try
            {
                newParams.lstExclude = new List<string>();
                newParams.lstExcludePrefix = new List<string>();
                newParams.lstExcludeSuffix = new List<string>();

                TextReader trConfigFile = new StreamReader("configComputerVerifier.txt");

                using (trConfigFile)
                {
                    string strNewLine = "";

                    while ((strNewLine = trConfigFile.ReadLine()) != null)
                    {

                        if (strNewLine.StartsWith("DisabledComputersLocation=") & strNewLine != "DisabledComputersLocation=")
                        {
                            newParams.strDisabledComputersLocation = strNewLine.Substring(26);
                            //[DebugLine] Console.WriteLine(newParams.strDisabledComputersLocation);
                        }
                        if (strNewLine.StartsWith("ExcludePrefix=") & strNewLine != "ExcludePrefix=")
                        {
                            newParams.lstExcludePrefix.Add(strNewLine.Substring(14).ToUpper());
                            //[DebugLine] Console.WriteLine(strNewLine.Substring(14));
                        }
                        if (strNewLine.StartsWith("ExcludeSuffix=") & strNewLine != "ExcludeSuffix=")
                        {
                            newParams.lstExcludeSuffix.Add(strNewLine.Substring(14).ToUpper());
                            //[DebugLine] Console.WriteLine(strNewLine.Substring(14));
                        }
                        if (strNewLine.StartsWith("Exclude=") & strNewLine != "Exclude=")
                        {
                            newParams.lstExclude.Add(strNewLine.Substring(8).ToUpper());
                            //[DebugLine] Console.WriteLine(strNewLine.Substring(8));
                        }
                        if (strNewLine.StartsWith("DaysSinceLastLogon=") & strNewLine != "DaysSinceLastLogon=")
                        {
                            newParams.intDaysSinceLastLogon = Int32.Parse(strNewLine.Substring(19));
                            //[DebugLine] Console.WriteLine(strNewLine.Substring(19) + newParams.intDaysSinceLastLogon.ToString());
                        }
                        if (strNewLine.StartsWith("DaysSinceDateCreated=") & strNewLine != "DaysSinceDateCreated=")
                        {
                            newParams.intDaysSinceDateCreated = Int32.Parse(strNewLine.Substring(21));
                            //[DebugLine] Console.WriteLine(strNewLine.Substring(21) + newParams.intDaysSinceDateCreated.ToString());
                        }
                        if (strNewLine.StartsWith("DisableUnusedComputers=") & strNewLine != "DisableUnusedComputers=")
                        {
                            newParams.strDisableUnusedComputers = strNewLine.Substring(23);
                            //[DebugLine] Console.WriteLine(newParams.strDisableUnusedComputers);
                        }
                    }
                }

                //[DebugLine] Console.WriteLine("# of Exclude= : {0}", newParams.lstExclude.Count.ToString());
                //[DebugLine] Console.WriteLine("# of ExcludePrefix= : {0}", newParams.lstExcludePrefix.Count.ToString());

                // Automatic account exclusions

                trConfigFile.Close();

                DateTime dtLastLogonFilter = DateTime.Today.AddDays(-newParams.intDaysSinceLastLogon);

                // [ Comment] Search filter strings for DirectorySearcher object filter
                string strFilterAll = "(&(&(objectCategory=computer)(name=*)" +
                                      "(!userAccountControl:1.2.840.113556.1.4.803:=2)(lastLogonTimestamp<=" +
                                      dtLastLogonFilter.ToFileTime().ToString() + ")))";

                string strFilterAllServers = "(&(&(&(&(sAMAccountType=805306369)(objectCategory=computer)" +
                                      "(!userAccountControl:1.2.840.113556.1.4.803:=2)(lastLogonTimestamp<=" +
                                      dtLastLogonFilter.ToFileTime().ToString() + ")" +
                                      "(|(operatingSystem=Windows Server 2008*)(operatingSystem=Windows Server 2003*)(operatingSystem=Windows 2000 Server*)(operatingSystem=Windows NT*)(operatingSystem=*2008*))))))";

                string strFilterAllWorkstations = "(&(&(&(&(sAMAccountType=805306369)(objectCategory=computer)" +
                                      "(!userAccountControl:1.2.840.113556.1.4.803:=2)(lastLogonTimestamp<=" +
                                      dtLastLogonFilter.ToFileTime().ToString() + ")" +
                                      "(|(operatingSystem=Windows XP Pro*)(operatingSystem=Windows 7*)(operatingSystem=Windows Vista*))))))";

                if (objCMDArguments2.strQueryFilter == "-all")
                {
                    newParams.strFilter = strFilterAll;
                }
                if (objCMDArguments2.strQueryFilter == "-allservers")
                {
                    newParams.strFilter = strFilterAllServers;
                }
                if (objCMDArguments2.strQueryFilter == "-allworkstations")
                {
                    newParams.strFilter = strFilterAllWorkstations;
                }

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

            return newParams;
        }

        static void funcProgramExecution(CMDArguments objCMDArguments2)
        {
            try
            {
                funcLogToEventLog("ComputerVerifier", "ComputerVerifier started", 1001);

                funcProgramRegistryTag("ComputerVerifier");

                ComputerLoginParams newParams = funcParseConfigFile(objCMDArguments2);

                funcCheckComputerLogin(newParams);

                funcLogToEventLog("ComputerVerifier", "ComputerVerifier stopped", 1002);
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcCheckComputerLogin(ComputerLoginParams newParams)
        {
            try
            {
                DateTime dtFilter = DateTime.Today.AddDays(-newParams.intDaysSinceLastLogon);

                // [Comment] Get local domain context
                string strrootDSE;

                System.DirectoryServices.DirectorySearcher objrootDSESearcher = new System.DirectoryServices.DirectorySearcher();
                strrootDSE = objrootDSESearcher.SearchRoot.Path;
                // [DebugLine]Console.WriteLine(rootDSE);

                // [Comment] Construct DirectorySearcher object using rootDSE string
                System.DirectoryServices.DirectoryEntry objrootDSEentry = new System.DirectoryServices.DirectoryEntry(strrootDSE);
                System.DirectoryServices.DirectorySearcher objComputerObjectSearcher = new System.DirectoryServices.DirectorySearcher(objrootDSEentry);
                // [DebugLine]Console.WriteLine(objComputerObjectSearcher.SearchRoot.Path);

                // [Comment] Add filter to DirectorySearcher object
                objComputerObjectSearcher.Filter = (newParams.strFilter);

                // [Comment] Execute query, return results, display name and path values
                System.DirectoryServices.SearchResultCollection objComputerResults = objComputerObjectSearcher.FindAll();
                // [DebugLine]Console.WriteLine(objComputerResults.Count.ToString());

                if(objComputerResults.Count > 0)
                {
                    funcLogToEventLog("ComputerVerifier", "Number of computers to process: " + objComputerResults.Count.ToString(), 1003);

                    TextWriter twCurrent = funcOpenOutputLog();
                    string strOutputMsg = "";

                    twCurrent.WriteLine("Date\tMessage");

                    string strlastLogonTimestamp = ""; //lastLogonTimestamp attribute
                    string strLastAccountLogon = "";
                    string strWhenCreated = "";
                    bool bValidLogonDate = false;

                    foreach (System.DirectoryServices.SearchResult objComputer in objComputerResults)
                    {
                        try
                        {
                            strOutputMsg = ""; //reset
                            strlastLogonTimestamp = ""; //reset
                            strLastAccountLogon = ""; //reset
                            strWhenCreated = ""; //reset
                            bValidLogonDate = false; //reset

                            System.DirectoryServices.DirectoryEntry objComputerDE = new System.DirectoryServices.DirectoryEntry(objComputer.Path);

                            if (!funcCheckNameExclusion(objComputerDE.Name.Substring(3), newParams))
                            {
                                Console.WriteLine("Computer: " + objComputerDE.Name.Substring(3));
                                strOutputMsg = "Computer: " + objComputerDE.Name.Substring(3);
                                funcWriteToOutputLog(twCurrent, strOutputMsg);

                                Console.WriteLine("Computer AD Path: " + objComputerDE.Path.Substring(7));
                                strOutputMsg = "Computer AD Path: " + objComputerDE.Path.Substring(7);
                                funcWriteToOutputLog(twCurrent, strOutputMsg);

                                strlastLogonTimestamp = funcGetLastLogonTimestamp(objComputerDE); // check for "(null)" or ""
                                strWhenCreated = funcGetAccountCreationDate(objComputerDE); // check for "(null)" or ""

                                //[DebugLine] Console.WriteLine("{0} \t {1} \t {2}", userDE.Name, strlastLogonTimestamp, strWhenCreated);

                                if (strlastLogonTimestamp != "(null)" & strlastLogonTimestamp != "")
                                {
                                    strLastAccountLogon = strlastLogonTimestamp;
                                    bValidLogonDate = true;
                                }
                                //else
                                //{
                                //    if (strlastLogonTimestamp == "(null)" | strlastLogonTimestamp == "")
                                //    {
                                //        if (u.LastLogon != null)
                                //        {
                                //            strLastAccountLogon = u.LastLogon.Value.ToLocalTime().ToString();
                                //            bValidLogonDate = true;
                                //        }
                                //    }
                                //}

                                if (bValidLogonDate)
                                {
                                    DateTime dtLastAccountLogon = Convert.ToDateTime(strLastAccountLogon);

                                    strOutputMsg = "";

                                    if (dtLastAccountLogon > dtFilter)
                                    {
                                        strOutputMsg = "Last login: " + objComputerDE.Name.Substring(3) + " - " + strLastAccountLogon;
                                    }
                                    else
                                    {
                                        int val = (int)objComputerDE.Properties["userAccountControl"].Value;
                                        objComputerDE.Properties["userAccountControl"].Value = val | 0x2;
                                        //ADS_UF_ACCOUNTDISABLE;
                                        objComputerDE.CommitChanges();
                                        //objComputerDE.Close();

                                        if (!funcIsDEActive(objComputerDE))
                                        {
                                            strOutputMsg = "Last login: " + objComputerDE.Name.Substring(3) + " - " + strLastAccountLogon + "\t(Action: Disabled)";
                                        }
                                        else
                                        {
                                            strOutputMsg = "Last login: " + objComputerDE.Name.Substring(3) + " - " + strLastAccountLogon + "\t(Action: NotDisabled-Check computer)";
                                        }
                                    }

                                    funcWriteToOutputLog(twCurrent, strOutputMsg);
                                }
                                else
                                {
                                    DateTime dtWhenCreated = Convert.ToDateTime(strWhenCreated);
                                    //[DebugLine] Console.WriteLine(dtWhenCreated.ToString());

                                    DateTime dtWhenCreatedCutOff = DateTime.Today.AddDays(-newParams.intDaysSinceDateCreated);

                                    strOutputMsg = "";

                                    if (dtWhenCreated < dtWhenCreatedCutOff)
                                    {
                                        if (newParams.strDisableUnusedComputers == "yes")
                                        {
                                            int val = (int)objComputerDE.Properties["userAccountControl"].Value;
                                            objComputerDE.Properties["userAccountControl"].Value = val | 0x2;
                                            //ADS_UF_ACCOUNTDISABLE;
                                            objComputerDE.CommitChanges();
                                            //objComputerDE.Close();

                                            if (!funcIsDEActive(objComputerDE))
                                            {
                                                strOutputMsg = "Unused computer: " + objComputerDE.Name.Substring(3) + "\t(Action: Disabled)";
                                            }
                                            else
                                            {
                                                strOutputMsg = "Unused computer: " + objComputerDE.Name.Substring(3) + "\t(Action: NotDisabled-Check computer)";
                                            }
                                        }
                                        else
                                        {
                                            strOutputMsg = "Unused computer: " + objComputerDE.Name.Substring(3) + "\t(NoAction: Outside allowed no-use period)";
                                        }
                                    }
                                    else
                                    {
                                        strOutputMsg = "Unused computer: " + objComputerDE.Name.Substring(3);
                                    }

                                    funcWriteToOutputLog(twCurrent, strOutputMsg);
                                }

                                bool bContactResult = funcContactServer(objComputerDE.Name.Substring(3, objComputerDE.Name.Length - 3));

                                if (bContactResult)
                                {
                                    Console.WriteLine("Contact was successful for: {0}", objComputerDE.Name.Substring(3));
                                    strOutputMsg = "Contact was successful for: " + objComputerDE.Name.Substring(3);
                                    funcWriteToOutputLog(twCurrent, strOutputMsg);
                                }
                                else
                                {
                                    Console.WriteLine("Contact was NOT successful for: {0}", objComputerDE.Name.Substring(3));
                                    strOutputMsg = "Contact was NOT successful for: " + objComputerDE.Name.Substring(3);
                                    funcWriteToOutputLog(twCurrent, strOutputMsg);
                                }
                            }
                        }
                        catch (SystemException ex)
                        {
                            // [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());
                            string strDirectoryServicesOperation = "There is no such object on the server.";
                            // [DebugLine] System.Console.WriteLine(ex.Message);
                            if (ex.Message == strDirectoryServicesOperation)
                            {
                                Console.WriteLine("An AD operation was unable to complete. Computer object was moved or deleted.");
                            }
                        }

                        Console.WriteLine();

                        funcCloseOutputLog(twCurrent);
                    }
                }

                objComputerResults.Dispose();              
            }
            catch
            {
            }
        }

        static bool funcCheckNameExclusion(string strName, ComputerLoginParams listParams)
        {
            try
            {
                bool bNameExclusionCheck = false;

                //List<string> listExclude = new List<string>();
                //listExclude.Add("Guest");
                //listExclude.Add("SUPPORT_388945a0");
                //listExclude.Add("krbtgt");

                strName = strName.ToUpper();

                if (listParams.lstExclude.Contains(strName))
                    bNameExclusionCheck = true;

                //string strMatch = listExclude.Find(strName);
                foreach (string strNameTemp in listParams.lstExcludePrefix)
                {
                    if (strName.StartsWith(strNameTemp))
                    {
                        bNameExclusionCheck = true;
                        break;
                    }
                }

                foreach (string strNameTemp in listParams.lstExcludeSuffix)
                {
                    if (strName.EndsWith(strNameTemp))
                    {
                        bNameExclusionCheck = true;
                        break;
                    }
                }

                return bNameExclusionCheck;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static void funcGetFuncCatchCode(string strFunctionName, Exception currentex)
        {
            string strCatchCode = "";

            Dictionary<string, string> dCatchTable = new Dictionary<string, string>();
            dCatchTable.Add("funcGetFuncCatchCode", "f0");
            dCatchTable.Add("funcPrintParameterWarning", "f2");
            dCatchTable.Add("funcPrintParameterSyntax", "f3");
            dCatchTable.Add("funcParseCmdArguments", "f4");
            dCatchTable.Add("funcProgramExecution", "f5");
            dCatchTable.Add("funcProgramRegistryTag", "f6");
            dCatchTable.Add("funcCreateDSSearcher", "f7");
            dCatchTable.Add("funcCreatePrincipalContext", "f8");
            dCatchTable.Add("funcCheckNameExclusion", "f9");
            dCatchTable.Add("funcMoveDisabledAccounts", "f10");
            dCatchTable.Add("funcFindAccountsToDisable", "f11");
            dCatchTable.Add("funcCheckLastLogin", "f12");
            dCatchTable.Add("funcRemoveUserFromGroup", "f13");
            dCatchTable.Add("funcToEventLog", "f14");
            dCatchTable.Add("funcCheckForFile", "f15");
            dCatchTable.Add("funcCheckForOU", "f16");
            dCatchTable.Add("funcWriteToErrorLog", "f17");

            if (dCatchTable.ContainsKey(strFunctionName))
            {
                strCatchCode = "err" + dCatchTable[strFunctionName] + ": ";
            }

            //[DebugLine] Console.WriteLine(strCatchCode + currentex.GetType().ToString());
            //[DebugLine] Console.WriteLine(strCatchCode + currentex.Message);

            funcWriteToErrorLog(strCatchCode + currentex.GetType().ToString());
            funcWriteToErrorLog(strCatchCode + currentex.Message);

        }

        static void funcWriteToErrorLog(string strErrorMessage)
        {
            try
            {
                string strPath = Directory.GetCurrentDirectory();

                if (!Directory.Exists(strPath + "\\Log"))
                {
                    Directory.CreateDirectory(strPath + "\\Log");
                    if (Directory.Exists(strPath + "\\Log"))
                    {
                        strPath = strPath + "\\Log";
                    }
                }
                else
                {
                    strPath = strPath + "\\Log";
                }

                FileStream newFileStream = new FileStream(strPath + "\\Err-ComputerVerifier.log", FileMode.Append, FileAccess.Write);
                TextWriter twErrorLog = new StreamWriter(newFileStream);

                DateTime dtNow = DateTime.Now;

                string dtFormat = "MMddyyyy HH:mm:ss";

                twErrorLog.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat), strErrorMessage);

                twErrorLog.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

        }

        static bool funcCheckForOU(string strOUPath)
        {
            try
            {
                string strDEPath = "";

                if (!strOUPath.Contains("LDAP://"))
                {
                    strDEPath = "LDAP://" + strOUPath;
                }
                else
                {
                    strDEPath = strOUPath;
                }

                if (DirectoryEntry.Exists(strDEPath))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static bool funcCheckForFile(string strInputFileName)
        {
            try
            {
                if (System.IO.File.Exists(strInputFileName))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static TextWriter funcOpenOutputLog()
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                string dtFormat2 = "MMddyyyy"; // for log file directory creation

                string strPath = Directory.GetCurrentDirectory();

                if (!Directory.Exists(strPath + "\\Log"))
                {
                    Directory.CreateDirectory(strPath + "\\Log");
                    if (Directory.Exists(strPath + "\\Log"))
                    {
                        strPath = strPath + "\\Log";
                    }
                }
                else
                {
                    strPath = strPath + "\\Log";
                }

                string strLogFileName = strPath + "\\ComputerVerifier" + dtNow.ToString(dtFormat2) + ".log";

                FileStream newFileStream = new FileStream(strLogFileName, FileMode.Append, FileAccess.Write);
                TextWriter twOuputLog = new StreamWriter(newFileStream);

                return twOuputLog;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return null;
            }

        }

        static void funcWriteToOutputLog(TextWriter twCurrent, string strOutputMessage)
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                string dtFormat = "MM/dd/yyyy";
                //string dtFormat2 = "MMddyyyy HH:mm:ss";

                twCurrent.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat), strOutputMessage);
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcCloseOutputLog(TextWriter twCurrent)
        {
            try
            {
                twCurrent.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static string funcGetLastLogonTimestamp(DirectoryEntry tmpDE)
        {
            string strTimestamp = "(null)";

            if (tmpDE.Properties.Contains("lastLogonTimestamp"))
            {
                //[DebugLine] Console.WriteLine(u.Name + " has lastLogonTimestamp attribute");
                IADsLargeInteger lintLogonTimestamp = (IADsLargeInteger)tmpDE.Properties["lastLogonTimestamp"].Value;
                if (lintLogonTimestamp != null)
                {
                    DateTime dtLastLogonTimestamp = funcGetDateTimeFromLargeInteger(lintLogonTimestamp);
                    if (dtLastLogonTimestamp != null)
                    {
                        strTimestamp = dtLastLogonTimestamp.ToString();
                    }
                    else
                    {
                        strTimestamp = "(null)";
                    }
                }
            }

            return strTimestamp;
        }

        static string funcGetAccountCreationDate(DirectoryEntry tmpDE)
        {
            string strCreationDate = "(null)";

            if (tmpDE.Properties.Contains("whenCreated"))
            {
                strCreationDate = tmpDE.Properties["whenCreated"].Value.ToString();
            }

            return strCreationDate;
        }

        static DateTime funcGetDateTimeFromLargeInteger(IADsLargeInteger largeIntValue)
        {
            //
            // Convert large integer to int64 value
            //
            long int64Value = (long)((uint)largeIntValue.LowPart +
                     (((long)largeIntValue.HighPart) << 32));

            //
            // Return the DateTime in utc
            //
            // return DateTime.FromFileTimeUtc(int64Value);


            // return in Localtime
            return DateTime.FromFileTime(int64Value);
        }

        static bool funcIsDEActive(DirectoryEntry de)
        {
            if (de.NativeGuid == null) return false;

            int flags = (int)de.Properties["userAccountControl"].Value;

            if (!Convert.ToBoolean(flags & 0x0002)) return true; else return false;
        }


        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    funcPrintParameterWarning();
                }
                else
                {
                    if (args[0] == "-?")
                    {
                        funcPrintParameterSyntax();
                    }
                    else
                    {
                        string[] arrArgs = args;
                        CMDArguments objArgumentsProcessed = funcParseCmdArguments(arrArgs);

                        if (objArgumentsProcessed.bParseCmdArguments)
                        {
                            funcProgramExecution(objArgumentsProcessed);
                        }
                        else
                        {
                            funcPrintParameterWarning();
                        } // check objArgumentsProcessed.bParseCmdArguments
                    } // check args[0] = "-?"
                } // check args.Length == 0
            }
            catch (Exception ex)
            {
                Console.WriteLine("errm0: {0}", ex.Message);
            }
        }
    }
}
