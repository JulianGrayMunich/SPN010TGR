using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Diagnostics;
using Microsoft.Data.SqlClient;

using databaseAPI;

using EASendMail;

using GNAgeneraltools;

using OfficeOpenXml;

using GNAspreadsheettools;

using GNAsurveytools;
using gnaDataClasses;

using Twilio.TwiML.Messaging;
using Twilio.Rest.Api.V2010.Account;
using Twilio.Rest.Sync.V1.Service.SyncStream;
using System.Reflection.Metadata.Ecma335;
using Twilio.TwiML.Voice;







namespace SPN010
{
    class Program
    {
        static void Main()
        {
#pragma warning disable CS0162
#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8604




            //================[Console settings]======================================
            Console.OutputEncoding = System.Text.Encoding.Unicode;

            //================[Declare variables]=====================================

            String[] strRO1 = new String[50];
            String[] strWorksheetName = new String[50];
            string[] strTrackWorksheets = new String[50];
            string[] strfudgeTargets = new String[20];

            //================[Configuration variables]==================================================================

            string strDBconnection = ConfigurationManager.ConnectionStrings["DBconnectionString"].ConnectionString;

            var config = ConfigurationManager.AppSettings;
            string strFreezeScreen = config["freezeScreen"];
            string strAlarmVersion = config["AlarmVersion"];
            string strDeleteMissingValues = config["DeleteMissingValues"];
            string strLatestValueOnly = config["LatestValueOnly"];

            string strProjectTitle = config["ProjectTitle"];
            string strContractTitle = config["ContractTitle"];
            string strReportType = config["ReportType"];
            string strReportSpec = config["ReportSpec"];

            string strExcelPath = config["ExcelPath"];
            string strExcelFile = config["ExcelFile"];
            string strCoordinateOrder = config["CoordinateOrder"];

            string strReferenceWorksheet = config["ReferenceWorksheet"];
            string strSurveyWorksheet = config["SurveyWorksheet"];
            string strCalibrationWorksheet = config["CalibrationWorksheet"];

            string strIncludeHistoricTwist= config["includeHistoricTwist"]; 


            string strSystemLogsFolder = config["SystemStatusFolder"];
            string strAlarmfolder = config["SystemAlarmFolder"];

            strTrackWorksheets[0] = strReferenceWorksheet;
            strTrackWorksheets[1] = config["Worksheet1"];
            strTrackWorksheets[2] = config["Worksheet2"];
            strTrackWorksheets[3] = config["Worksheet3"];
            strTrackWorksheets[4] = config["Worksheet4"];
            strTrackWorksheets[5] = config["Worksheet5"];
            strTrackWorksheets[6] = config["Worksheet6"];
            strTrackWorksheets[7] = config["Worksheet7"];
            strTrackWorksheets[8] = config["Worksheet8"];
            strTrackWorksheets[9] = config["Worksheet9"];
            strTrackWorksheets[10] = config["Worksheet10"];
            strTrackWorksheets[11] = "blank";


            strfudgeTargets[0] = "blank";
            strfudgeTargets[1] = config["fudgeTarget1"];
            strfudgeTargets[2] = config["fudgeTarget2"];
            strfudgeTargets[3] = config["fudgeTarget3"];
            strfudgeTargets[4] = config["fudgeTarget4"];
            strfudgeTargets[5] = config["fudgeTarget5"];
            strfudgeTargets[6] = config["fudgeTarget6"];
            strfudgeTargets[7] = config["fudgeTarget7"];
            strfudgeTargets[8] = config["fudgeTarget8"];
            strfudgeTargets[9] = config["fudgeTarget9"];
            strfudgeTargets[10] = config["fudgeTarget10"];
            strfudgeTargets[11] = "blank";

            string strFirstDataRow = config["FirstDataRow"];
            string strFirstOutputRow = config["FirstOutputRow"];
            string strFirstDataCol = config["FirstDataCol"];

            string strTimeBlockType = config["TimeBlockType"];
            string strStartTimeOffset = config["StartTimeOffsetHrs"];
            string strDataBlockSize = config["BlockSizeHrs"];
            string strManualTimeBlockStart = config["manualTimeBlockStart"];
            string strManualBlockStart = config["manualBlockStart"];
            string strManualBlockEnd = config["manualBlockEnd"];
            string strTimeOffsetHrs = config["TimeOffsetHrs"];
            string strBlockSizeHrs = config["BlockSizeHrs"];


            string strTimeBlockStartLocal = "";
            string strTimeBlockEndLocal = "";
            string strTimeBlockStartUTC = "";
            string strTimeBlockEndUTC = "";
            string strEmailTime = "";

            string strTempString = "";

            string strSPN010alarms = config["SPN010alarmNotifications"];
            string strSMSTitle = config["SMSTitle"];

            int iRow = Convert.ToInt32(strFirstDataRow);
            int iReferenceFirstDataRow = Convert.ToInt32(strFirstDataRow);
            int iFirstOutputRow = Convert.ToInt32(strFirstOutputRow);
            int iCol = Convert.ToInt32(strFirstDataCol);


            string strSendEmails = config["SendEmails"];
            string strEmailLogin = config["EmailLogin"];
            string strEmailPassword = config["EmailPassword"];
            string strEmailFrom = config["EmailFrom"];
            string strEmailRecipients = config["EmailRecipients"];


            string strMasterWorkbookFullPath = strExcelPath + strExcelFile;
            string[,] strSensorID = new string[5000, 2];
            string[,] strPointDeltas = new string[5000, 2];
            string strDateTime = "";
            string strMasterFile = "";
            string strWorkingFile = "";
            string strExportFile = "";


            #region SMS numbers

            List<string> smsMobile = new();
            string strMobileList = "";
            var allKeys = config.AllKeys;
            var recipientKeys = allKeys.Where(k => k != null && k.StartsWith("RecipientPhone"));

            foreach (string key1 in recipientKeys)
            {
                string value = config[key1];
                if (!string.IsNullOrWhiteSpace(value))
                {
                    smsMobile.Add(value);
                    if (strMobileList != "") strMobileList += ",";
                    strMobileList += value;
                }
            }
            //Console.WriteLine(strTab1 + "Mobile list: " + strMobileList);
            //Console.ReadKey();
            #endregion

            string strTab1 = "     ";
            string strTab2 = "        ";
            string strTab3 = "           ";


            //================[Actions]======================================
            // 20240426 Creation of email & SMS alarm facility
            // 20241014 Updated email function
            // 20241015 basic housekeeping before big changes
            // 20241018 Add historic twist worksheet
            // 20241019 Add fudge values
            // 20241029 Add ability to take means instead of just latest value
            // 20241118 Updated mean/latest message as there seems to be an issue with this.
            // 20241119 Install filter to weed out spurious readings in writeLatestDeltas

            //================[Main program]===========================================================================

            //==== Set the EPPlus license
            ExcelPackage.License.SetCommercial("14XO1NhmOmVcqDWhA0elxM72um6vnYOS8UiExVFROZuRPn1Ddv5fRV8fiCPcjujkdw9H18nExINNFc8nmOjRIQEGQzVDRjMz5wdPAJkEAQEA");  //valid to 23.03.2026


            // instantiate the classes

            gnaTools gnaT = new();
            dbAPI gnaDBAPI = new();
            GNAsurveycalcs gnaSurvey = new();
            spreadsheetAPI gnaSpreadsheetAPI = new();
            gnaDataClass gnaDC = new();

            // Welcome message
            gnaT.WelcomeMessage("SPN010TGR 20250902");

            ExcelPackage.License.SetCommercial("14XO1NhmOmVcqDWhA0elxM72um6vnYOS8UiExVFROZuRPn1Ddv5fRV8fiCPcjujkdw9H18nExINNFc8nmOjRIQEGQzVDRjMz5wdPAJkEAQEA");  //valid to 23.03.2026

            //==== Environment check


            Console.WriteLine("");
            Console.WriteLine("1. Check system environment");


            gnaDBAPI.testDBconnection(strDBconnection);

            //==== Environment check
            Console.WriteLine("\n1. Check system environment");
            Console.WriteLine(strTab1 + "Check DB connection");
            gnaDBAPI.testDBconnection(strDBconnection);
            Console.WriteLine(strTab2 + "Done");

            Console.WriteLine(strTab1 + "Check existence of workbook & worksheets");
            if (strFreezeScreen == "Yes")
            {
                Console.WriteLine(strTab2 + "Project: " + strProjectTitle);
                Console.WriteLine(strTab2 + "Report type: " + strReportSpec);
                Console.WriteLine(strTab2 + "Master workbook: " + strMasterWorkbookFullPath);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strReferenceWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strSurveyWorksheet);
                int i = 1;
                do
                {
                    string strTrackWorksheet = strTrackWorksheets[i].Trim();
                    gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strTrackWorksheet);

                    if (strIncludeHistoricTwist == "Yes"){
                        gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strTrackWorksheet+"_HistoricTwist");
                    }
                    i++;
                } while (strTrackWorksheets[i] != "blank");

                Console.WriteLine(strTab2 + "Done");
            }
            else
            {
                Console.WriteLine(strTab2 + "Workbook & worksheets not checked");
            }



//==== Prepare the time block

            switch (strTimeBlockType)
            {
                case "Manual":
                    strTimeBlockStartLocal = strManualBlockStart;
                    strTimeBlockEndLocal = strManualBlockEnd;
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strTimeBlockStartLocal);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strTimeBlockEndLocal);
                    strEmailTime = string.Concat(strTimeBlockEndLocal.Replace("'", ""), "m");

                    break;

                case "Schedule":

                    //double dblStartTimeOffset = -1.0 * Convert.ToDouble(strTimeOffsetHrs);
                    double dblEndTimeOffset = -1.0 * Convert.ToDouble(strBlockSizeHrs);
                    strTimeBlockEndLocal = " '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockStartLocal = " '" + DateTime.Now.AddHours(dblEndTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strTimeBlockStartLocal);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strTimeBlockEndLocal);
                    break;
                default:
                    Console.WriteLine("\nError in Timeblock Type");
                    Console.WriteLine(strTab1 + "Time block type: " + strTimeBlockType);
                    Console.WriteLine(strTab1 + "Must be Manual or Schedule");
                    Console.WriteLine("\nPress key to exit..."); Console.ReadKey();
                    goto ThatsAllFolks;
                    break;
            }

            strDateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");
            string strDateTimeUTC = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm");   //2022-07-26 13:45:15
            string strTimeStamp = strTimeBlockEndLocal + "\n(local)";
            
            Console.WriteLine("\n" + strTab1 + "Time block type: " + strTimeBlockType);
            Console.WriteLine(strTab2 + strTimeBlockStartLocal.Replace("'","") + " Local");
            Console.WriteLine(strTab2 + strTimeBlockEndLocal.Replace("'", "") + " Local\n");

            string strTimeStampLocal = "";

            if (strTimeBlockType == "Manual")
            {
                string strTemp = strEmailTime.Replace(":", "").Replace("-", "").Replace(" ", "_");
                strExportFile = strExcelPath + strContractTitle + "_" + strReportType + "_" + strTemp + ".xlsx";
                strWorkingFile = strExportFile;
                strMasterFile = strExcelPath + strExcelFile;
                strTimeStampLocal = strTemp;
            }
            else
            {
                strExportFile = strExcelPath + strContractTitle + "_" + strReportType + "_" + strDateTime + ".xlsx";
                strWorkingFile = strExportFile;
                strMasterFile = strExcelPath + strExcelFile;
                strTimeStampLocal = strDateTime;
            }

            

            //==== Process data ===================================================================================

            Console.WriteLine("2. Extract point names");
            string[] strPointNames = gnaSpreadsheetAPI.readPointNames(strMasterFile, strSurveyWorksheet, strFirstDataRow);
            Console.WriteLine(strTab1 + "Done");
            Console.WriteLine("3. Extract SensorID");
            strSensorID = gnaDBAPI.getSensorIDfromDB(strDBconnection, strPointNames, strProjectTitle);
            Console.WriteLine(strTab1 + "Done");
            Console.WriteLine("4. Write SensorID to workbook");
            gnaSpreadsheetAPI.writeSensorID(strMasterFile, strSurveyWorksheet, strSensorID, strFirstDataRow);
            Console.WriteLine(strTab1 + "Done");

            if (strLatestValueOnly == "Yes")
            {
                Console.WriteLine("5. Extract latest deltas for time block");
                strPointDeltas = gnaDBAPI.getLatestDeltasFromDB(strDBconnection, strProjectTitle, strTimeBlockStartUTC, strTimeBlockEndUTC, strSensorID);
                strTempString = "latest";
                Console.WriteLine(strTab1 + "Done");

            }
            else
            {
                Console.WriteLine("5. Extract mean deltas for UTC time block");
                Console.WriteLine(strTab1 + strTimeBlockStartUTC.Replace("'", ""));
                Console.WriteLine(strTab1 + strTimeBlockEndUTC.Replace("'", ""));
                strPointDeltas = gnaDBAPI.getMeanDeltasFromDB(strDBconnection, strProjectTitle, strTimeBlockStartUTC, strTimeBlockEndUTC, strSensorID);
                strTempString = "mean";
                Console.WriteLine(strTab1 + "Done");
            }

            Console.WriteLine("6. Write " + strTempString+ " deltas & timestamp to master workbook");
            string strBlockStart = strTimeBlockStartUTC.Replace("'", "").Trim();
            string strBlockEnd = strTimeBlockEndUTC.Replace("'", "").Trim();
           
            gnaSpreadsheetAPI.writeLatestDeltas(
                strMasterFile, 
                strReferenceWorksheet, 
                strPointDeltas, 
                iRow, iCol, strBlockStart, 
                strBlockEnd, 
                strCoordinateOrder);

            gnaSpreadsheetAPI.writeTimeStampLocal(
                strMasterFile,
                strReferenceWorksheet,
                strTimeStampLocal);


            Console.WriteLine("7. Write historic twist for each line");

            // write the historic twist data if applicable
            if (strIncludeHistoricTwist=="Yes")
            {
                int iFirstEmptyCol = 0;
                int i = 1;
                do
                {
                    string strTrackWorksheet = strTrackWorksheets[i].Trim();
                    string strHistoricTwistWorksheet = strTrackWorksheet + "_HistoricTwist";
                    Console.WriteLine( strTab1+ strHistoricTwistWorksheet);
                    iFirstEmptyCol = gnaSpreadsheetAPI.findFirstEmptyColumn(strMasterFile, strHistoricTwistWorksheet, "6", "1");
                    int iSourceCol = iFirstEmptyCol - 1;
                    int iDestinationCol = iFirstEmptyCol;

                    gnaSpreadsheetAPI.copyColumnBetweenWorksheets(strMasterFile, strTrackWorksheet, strHistoricTwistWorksheet, 12, 6, iFirstEmptyCol, 6, strTimeBlockEndLocal, 1.0);
                    i++;
                } while (strTrackWorksheets[i] != "blank");

                gnaSpreadsheetAPI.formatHistoricTwist(strMasterFile, strReferenceWorksheet, strfudgeTargets, strTrackWorksheets);

                Console.WriteLine(strTab1 + "Done");

            }
            else
            {
                Console.WriteLine(strTab1 + "Not activated");
            }


            //Console.WriteLine("7. Calibration data");
            //string strDistanceColumn = "3";
            //gnaSpreadsheetAPI.populateCalibrationWorksheet(strDBconnection, strTimeBlockStartUTC, strTimeBlockEndUTC, strWorkingFile, strCalibrationWorksheet, strFirstOutputRow, strDistanceColumn, strProjectTitle);


            Console.WriteLine("8. Create the export workbook");
            gnaSpreadsheetAPI.copyWorkbook(strMasterFile, strExportFile);
            Console.WriteLine(strTab1+ strExportFile);
            Console.WriteLine(strTab1 + "Done");

            Console.WriteLine("9. Clean export workbook to match SPN010 template");
            int j = 1;
            do
            {
                string strTrackWorksheet = strTrackWorksheets[j].Trim();
                Console.WriteLine(strTab1 + strTrackWorksheet);
                // convert Columns 2 & 6 to numbers
                Console.WriteLine(strTab2 + "Convert references to values");
                gnaSpreadsheetAPI.convertWorksheetFormulae(strExportFile, strTrackWorksheet, iFirstOutputRow, 2, 2);    // Left rail reduced level at target
                gnaSpreadsheetAPI.convertWorksheetFormulae(strExportFile, strTrackWorksheet, iFirstOutputRow, 6, 6);    // Right rail prism ht

                if (strDeleteMissingValues == "Yes")
                {
                    Console.WriteLine(strTab2 + "Delete missing data");
                    gnaSpreadsheetAPI.removeSPN010missingData(strExportFile, strTrackWorksheet);
                    Console.WriteLine(strTab3 + "Done");
                }
                else
                {
                    Console.WriteLine(strTab2 + "Missing data not deleted");
                }

                j++;
            } while (strTrackWorksheets[j] != "blank");
            Console.WriteLine(strTab1 + "Done");


            //EnterHere:
            Console.WriteLine("10. Check alarm state");

            
            string txtMessage = "";

            //for (j = 1; j <= 9; j++)
            //{
            //    if (smsMobile[j] != "None")
            //    {
            //        strMobileList = strMobileList + smsMobile[j] + ",";
            //    }
            //}
            //strMobileList = strMobileList.Substring(0, strMobileList.Length - 1);

            string smsAlarmState = "No alarms";
            if (strSPN010alarms == "Yes")
            {
                string smsMessage = "";
                string strFullSMSmessage = "";
                j = 1;
                do
                {
                    string strTrackWorksheet = strTrackWorksheets[j].Trim();
                    //strExportFile = "C:\\_Working_drive\\Woodbrook\\SPN010alarms\\AlarmTest.xlsx";

                    //strExportFile = "C:\\Woodbrook\\SPN010\\AlarmCheck\\AlarmTest.xlsx";

                    smsMessage = gnaSpreadsheetAPI.SPN010alarms(strExportFile, strTrackWorksheet, strSMSTitle);

                    if (smsMessage != "No alarms")
                    {
                        smsAlarmState = "Alarms";
                    }

                    strFullSMSmessage = strFullSMSmessage + smsMessage + "\n";
                    j++;
                } while (strTrackWorksheets[j] != "blank");

                string strCurrentAlarmState = strFullSMSmessage;

                // Check whether the alarm state has changed & update the Alarm log
                string strSMSaction = gnaT.updateSystemAlarmFile(strAlarmfolder, strCurrentAlarmState);
                //strSMSaction = "SendSMS" / "DoNotSendSMS";
                if (strSMSaction == "SendSMS")
                {
                    // Strip out the string "No alarms\n"
                    Console.WriteLine(strTab1 + "Send SMS");

                    strFullSMSmessage = strFullSMSmessage.Replace("No alarms\n", "").Trim();

                    if (strFullSMSmessage.Length == 0)
                    {
                        strFullSMSmessage = "alarms cancelled";
                    }

                    // Send the SMS 
                    gnaT.sendSMSArray(strFullSMSmessage, smsMobile);

                    bool smsSuccess = gnaT.sendSMSArray(strFullSMSmessage, smsMobile);
                    Console.WriteLine(strTab1 + (smsSuccess ? "SMS sent" : "SMS failed"));
                    string strMessage = "";
                    if (smsSuccess == true)
                    {
                        strMessage = "SPN010 Report: SMS message sent";
                    }
                    else
                    {
                        strMessage = "SPN010 Report: SMS message failed";
                    }

                    gnaT.updateSystemLogFile(strSystemLogsFolder, strMessage);








                    string strNow = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm") + " : ";
                    string strEmailSubstring = "";
                    if (smsAlarmState == "No alarms")
                    {
                        strEmailSubstring = strFullSMSmessage;
                    }
                    else
                    {
                        strEmailSubstring = "in alarm state";
                    }

                    strMessage = strNow + "SPN010 Alarm SMS: " + strSMSTitle + " " + strEmailSubstring + " (" + strMobileList + ")";
                    gnaT.updateSystemLogFile(strSystemLogsFolder, strMessage);

                }
                else
                {
                    Console.WriteLine(strTab1 + "No SMS sent");
                }

                if ((smsAlarmState == "Alarms") && (strSMSaction == "SendSMS"))
                {
                    smsAlarmState = "New alarms";
                    txtMessage = "Send SMS";
                }


                if ((smsAlarmState == "Alarms") && (strSMSaction == "DoNotSendSMS"))
                {
                    smsAlarmState = "Existing alarms";
                    txtMessage = "No SMS";

                }


                if ((smsAlarmState == "No alarms") && (strSMSaction == "SendSMS"))
                {
                    smsAlarmState = "Alarms cancelled";
                    txtMessage = "Send SMS";
                }

                if ((smsAlarmState == "No alarms") && (strSMSaction == "DoNotSendSMS"))
                {
                    smsAlarmState = "No alarms";
                    txtMessage = "No SMS";
                }

                Console.WriteLine(strTab1 + smsAlarmState);
                Console.WriteLine(strTab1 + txtMessage);

                //goto ThatsAllFolks;

                // email the activity

                if (strSMSaction == "SendSMS")
                {
                    Console.WriteLine(strTab1+"Send alarm email");
                    strDateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");
                    string strAlarmHeading = "ALARM STATUS:" + strSMSTitle + " (" + strDateTime + ")";
                    string strMessage = gnaT.addCopyright("SPN010", strFullSMSmessage);
                    // updated with the 20240816 license
                    string license = gnaT.commercialSoftwareLicense("email");
                    SmtpMail oMailSMS = new(license)
                    {
                        From = strEmailFrom,
                        To = new AddressCollection(strEmailRecipients),
                        Subject = strAlarmHeading,
                        TextBody = strFullSMSmessage
                    };

                    // SMTP server address
                    SmtpServer oServerSMS = new("smtp.gmail.com")
                    {
                        User = strEmailLogin,
                        Password = strEmailPassword,
                        ConnectType = SmtpConnectType.ConnectTryTLS,
                        Port = 587
                    };

                    //Set sender email address, please change it to yours


                    SmtpClient oSmtpSMS = new();
                    oSmtpSMS.SendMail(oServerSMS, oMailSMS);

                    strMessage = strAlarmHeading + " (emailed " + strEmailRecipients + ")";

                    gnaT.updateSystemLogFile(strSystemLogsFolder, strMessage);

                    Console.WriteLine(strTab1 + "Alarm email sent");
                }
                else
                {
                    Console.WriteLine(strTab1 + "No alarm email sent");

                    if (strAlarmVersion == "Yes")
                    {
                        // delete the working file
                        Console.WriteLine(strTab1 + "Working file deleted");
                        gnaSpreadsheetAPI.deleteWorkbook(strWorkingFile);
                        Console.WriteLine(strTab1 + "Export file deleted");
                        // delete the export file strExportFile
                        gnaSpreadsheetAPI.deleteWorkbook(strExportFile);
                    }

                }
            }

            Console.WriteLine("11. email the export workbook");

            if ((strSendEmails == "Yes") && (strAlarmVersion == "No"))
            {
               
                try
                {
                    string strMessage = "This is an automated " + strReportSpec + " track geometry report. \nPlease review and forward to the client. \nDo not reply to this email.";
                    strMessage = gnaT.addCopyright("SPN010", strMessage);

                    // updated with the 20240816 license
                    // updated with the 20240816 license
                    string license = gnaT.commercialSoftwareLicense("email");
                    SmtpMail oMailEmail = new(license)
                    {
                        //Set sender email address, please change it to yours
                        From = strEmailFrom,
                        To = new AddressCollection(strEmailRecipients),
                        Subject = "SPN010: " + strProjectTitle + " (" + strDateTime + ")",
                        TextBody = strMessage
                    };
                    oMailEmail.AddAttachment(strExportFile);
                    // SMTP server address
                    SmtpServer oServerEmail = new("smtp.gmail.com")
                    {
                        User = strEmailLogin,
                        Password = strEmailPassword,
                        ConnectType = SmtpConnectType.ConnectTryTLS,
                        Port = 587
                    };

                    //Set sender email address, please change it to yours
                    SmtpClient oSmtpEmail = new();
                    oSmtpEmail.SendMail(oServerEmail, oMailEmail);
                    strMessage = strReportSpec + " Track Geometry Report: " + strProjectTitle + " (" + strDateTime + ")" + " (emailed)";

                    gnaT.updateSystemLogFile(strSystemLogsFolder, strMessage);
                    gnaT.updateReportTime("SPN010");

                    Console.WriteLine(strTab1+"Done");

                }
                catch (Exception ep)
                {
                    Console.WriteLine("Failed to send email with the following error:");
                    Console.WriteLine(strEmailLogin);
                    Console.WriteLine(strEmailPassword);
                    Console.WriteLine(ep.Message);
                    Console.ReadKey();
                }
            }
            else
            {
                Console.WriteLine(strTab1 + "No email sent");
            }

ThatsAllFolks:

            Console.WriteLine("\nSPN010 report completed...");
            gnaT.freezeScreen(strFreezeScreen);
            Environment.Exit(0);

        }
    }
}
