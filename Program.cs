using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.IO;
using System.Data;
using System.XLS;
using System.CSV;
using System.Collections;

// new entry
// ver 2
namespace Excel2CSV
{
    class Program
    {
        // return code enums
        public enum returnCodes
        {
            Success = 0,
            Config = 1,
            Parameter = 2,
            Customer = 4,
            Product = 8,
            Exception = 16
        }

        // global vars
        static string mLogFile;
        static DataSet dsData;
        static Dictionary<string, string> configDictionary;
        public static int mTotalFiles;
        public static int mFileCount;
        public static int mErrorCount;
        public static int mOuputCount;

        // site lookup dictionary
        public static Dictionary<string, string> dictSite8to7;


        static int Main(string[] args)
        {
            try
            {

                mLogFile = ConfigurationManager.AppSettings["ACTIVELOG"].ToString();
                Utils.WriteLog(mLogFile, "XLS2CSV program started...");

                // Read parameters from config file
                configDictionary = new Dictionary<string, string>();
                try
                {
                    configDictionary.Add("INPUTFOLDER", ConfigurationManager.AppSettings["INPUTFOLDER"].ToString());
                    configDictionary.Add("OUTPUTFOLDER", ConfigurationManager.AppSettings["OUTPUTFOLDER"].ToString());
                    configDictionary.Add("ARCHIVEFOLDER", ConfigurationManager.AppSettings["ARCHIVEFOLDER"].ToString());
                    configDictionary.Add("XLSRANGE", ConfigurationManager.AppSettings["XLSRANGE"].ToString());
                    configDictionary.Add("EMPOUTPUTCOLS", ConfigurationManager.AppSettings["EMPOUTPUTCOLS"].ToString());
                    configDictionary.Add("PAYOUTPUTCOLS", ConfigurationManager.AppSettings["PAYOUTPUTCOLS"].ToString());
                    configDictionary.Add("CSVHDRS", ConfigurationManager.AppSettings["CSVHDRS"].ToString());
                    configDictionary.Add("CSVDATEFMT", ConfigurationManager.AppSettings["CSVDATEFMT"].ToString());
                    configDictionary.Add("DELIMITER", ConfigurationManager.AppSettings["DELIMITER"].ToString());
                    configDictionary.Add("SITELIST", ConfigurationManager.AppSettings["SITELIST"].ToString());
                    configDictionary.Add("FLAGCOMPLETED", ConfigurationManager.AppSettings["FLAGCOMPLETED"].ToString());
                    configDictionary.Add("FLAGSUCCESS", ConfigurationManager.AppSettings["FLAGSUCCESS"].ToString());

                }
                catch (Exception ex)
                {
                    Utils.WriteLog(mLogFile, string.Concat("!ERROR parsing config file", ex.Message));
                    return (short)returnCodes.Config;
                }

                short iResult = (short)returnCodes.Success;

                // LOAD SITE DICTIONARY
                LoadSite8to7Dictionary(configDictionary["SITELIST"]);

                // process any files waiting
                ProcessAllFiles();

                if (mErrorCount==0)
                    Utils.WriteLog(configDictionary["FLAGSUCCESS"], iResult.ToString(), true);


                // write completed flag for Remoteware
                Utils.WriteLog(configDictionary["FLAGCOMPLETED"], string.Concat("Completed with exitcode:", iResult.ToString(), ". Files processed:", mFileCount.ToString(), ". Output Files:", mOuputCount.ToString()), true);
                Utils.WriteLog(mLogFile, string.Concat("Completed with exitcode:", iResult.ToString(), ". Files processed:", mFileCount.ToString(), ". Output Files:",mOuputCount.ToString()  ));


                return iResult;

            }
            catch (Exception ex)
            {
                Utils.WriteLog(mLogFile, string.Concat("!ERROR encountered:", ex.Message));
                Utils.WriteLog(configDictionary["FLAGCOMPLETED"], ex.Message, true);
                return (short)returnCodes.Exception;
            }

        }


        //--------------------------------------------------------------------------------------------------------
        public static void ProcessAllFiles()
        {

            string xlsPath = configDictionary["INPUTFOLDER"];
            
            short iResult;
            mFileCount = 0;
            mErrorCount = 0;
            mOuputCount = 0;
            string[] fileEntries = Directory.GetFiles(xlsPath, "*.xls", SearchOption.TopDirectoryOnly);
            mTotalFiles = fileEntries.Length;
            foreach (string xls in fileEntries)
            {
                // get date suffix
                string shortName = Path.GetFileNameWithoutExtension(xls);
                string[] aShortName = shortName.Split('_');
                string csvSuffix = aShortName[aShortName.Length - 1];

               
                // process file - could produce multiple output files
                iResult = ProcessXLS(xls, csvSuffix, configDictionary["XLSRANGE"]);

                if (iResult == 0)
                    MoveFile(xls, configDictionary["ARCHIVEFOLDER"]);
                else
                    mErrorCount++;

                // update progress
                mFileCount++;
                

            }

        }

        // --------------------------------------------------------------------------------------------
        private static short ProcessXLS(string fileName, string csvSuffix, string xlsRange)
        {

            string outputFile;
            string siteCode;
            string includeCols;

            if (!File.Exists(fileName))
            {
                Utils.WriteLog(mLogFile, string.Concat("Could not find ", fileName));
                return (0);
            }

            short iResult;

            try
            {

                // Process xls
                Utils.WriteLog(mLogFile, string.Concat("Processing ", fileName));
                XLS objXLS = new XLS();

                // output to CSV
                CSV objCSV = new CSV();

                // extract range to dataset
                dsData = objXLS.GetXLSData(fileName, xlsRange, true);

                if (dsData != null)
                {
                    // loop through each distinct location code in data
                    DataView dvSites = new DataView(dsData.Tables[0]);
                    DataTable dtSites = dvSites.ToTable(true, "Location");

                    foreach (DataRow row in dtSites.Rows)
                    {
                        string location = row["Location"].ToString();
                        // lookup 7 digit
                        siteCode = GetSite8Digit(location);

                        // filter data by site
                        DataView dvSingleSite = new DataView(dsData.Tables[0]);
                        dvSingleSite.RowFilter = string.Concat("Location='", location, "'");
                        DataTable dtSingleSite = dvSingleSite.ToTable();

                        // ----------------------------------------------------------------------
                        // create employee output
                        // ----------------------------------------------------------------------

                        // get new csv name
                        includeCols = configDictionary["EMPOUTPUTCOLS"];
                        outputFile = string.Concat(configDictionary["OUTPUTFOLDER"], @"\Employee", siteCode, csvSuffix, ".csv");

                        // create file for site
                        iResult = objCSV.CreateCSVFile(dtSingleSite, outputFile, includeCols, configDictionary["CSVDATEFMT"], configDictionary["CSVHDRS"], configDictionary["DELIMITER"], siteCode);
                        
                        
                        switch (iResult)
                        {
                            case 0:
                                Utils.WriteLog(mLogFile, string.Concat("Created output ", outputFile));
                                mOuputCount++;
                                break;
                            case -1:
                                Utils.WriteLog(mLogFile, "!ERROR! Expected columns not found creating ");
                                break;
                            default:
                                Utils.WriteLog(mLogFile, "!ERROR! Undefined error  ");
                                break;

                        }
                        // stop on error
                        if (iResult!=0) return (iResult);

                        // ----------------------------------------------------------------------
                        // create payrates output
                        // ----------------------------------------------------------------------

                        // get new csv name
                        includeCols = configDictionary["PAYOUTPUTCOLS"];
                        outputFile = string.Concat(configDictionary["OUTPUTFOLDER"], @"\PayRates", siteCode, csvSuffix, ".csv");

                        // create file for site
                        iResult = objCSV.CreateCSVFile(dtSingleSite, outputFile, includeCols, configDictionary["CSVDATEFMT"], configDictionary["CSVHDRS"], configDictionary["DELIMITER"], siteCode);


                        switch (iResult)
                        {
                            case 0:
                                Utils.WriteLog(mLogFile, string.Concat("Created output ", outputFile));
                                mOuputCount++;
                                break;
                            case -1:
                                Utils.WriteLog(mLogFile, "!ERROR! Expected columns not found creating ");
                                break;
                            default:
                                Utils.WriteLog(mLogFile, "!ERROR! Undefined error  ");
                                break;

                        }
                        // stop on error
                        if (iResult != 0) return (iResult);

                    }
                    
                }

                return (0);
            }
            catch (Exception ex)
            {
                Utils.WriteLog(mLogFile, string.Concat("!ERROR processing ", fileName, ex.Message));
                return -1;
            }

        }



        //--------------------------------------------------------------------------------------------------------
        private static void MoveFile(string fullName, string destPath)
        {
            string filename = Path.GetFileName(fullName);
            string destFile = string.Concat(destPath, @"\", filename);
            if (File.Exists(destFile)) File.Delete(destFile);
            try
            {
                File.Move(fullName, destFile);
            }
            catch
            {
                // continue 
                Utils.WriteLog(mLogFile, string.Concat("Could not move file to: ", destFile));
            }

        }


        // -------------------------------------------------------------------------------------------------------
        // Load site lookup file to dictionary
        // -------------------------------------------------------------------------------------------------------
        private static Boolean LoadSite8to7Dictionary(string siteList)
        {
            // check file exists
            if (!File.Exists(siteList))
            {
                Utils.WriteLog(mLogFile, string.Concat("Site lookup file not found:", siteList));
                return (false);
            }

            dictSite8to7 = new Dictionary<string, string>();
            foreach (string s in File.ReadAllLines(siteList))
            {
                var vals = s.Split('|');
                dictSite8to7[vals[1]] = vals[0];

            }
            return true;

        }

        // -------------------------------------------------------------------------------------------------------
        // Get 7 digit from 8 digit
        // -------------------------------------------------------------------------------------------------------
        private static string GetSite8Digit(string key)
        {
            try
            {
                return dictSite8to7[key];
            }
            catch
            {
                Utils.WriteLog(mLogFile, string.Concat("!ERROR could not find 8 digit sitecode for ", key));
                return "00000000";
            }


        }

    }
}
