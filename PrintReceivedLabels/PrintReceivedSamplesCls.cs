using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Entity;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using LSEXT;
using Patholab_Common;
using System.Linq;
using Patholab_DAL_V1;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using PrintGeneralPatholab;


namespace PrintReceivedLabels
{

    [ComVisible(true)]
    [ProgId("PrintReceivedLabels.PrintReceivedSamplesCls")]
    public class PrintReceivedSamplesCls : IWorkflowExtension //Labels for sample
    {
        INautilusServiceProvider sp;
        private const string Type = "3";
        private int _port = 9100;
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
        uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);
        private DataLayer dal;
        private AppSettingsSection _appSettings;
        private bool _debug;



        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {
           
                #region param

                string role = Parameters["ROLE_NAME"];


                _debug = (role.ToUpper() == "DEBUG");

                if (_debug) Debugger.Launch();

                string tableName = Parameters["TABLE_NAME"];
                sp = Parameters["SERVICE_PROVIDER"];
                var rs = Parameters["RECORDS"];
                rs.MoveLast();
                long tableID;
            
                 
                if (tableName != "SAMPLE")
                {
                    MessageBox.Show("This program works on SAMPLE only.");
                    return;
                }


                #endregion

                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                ExeConfigurationFileMap map = new ExeConfigurationFileMap();
                map.ExeConfigFilename = assemblyPath + ".config";
                Configuration cfg = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                _appSettings = cfg.AppSettings;
                string aaaa = _appSettings.Settings[tableName + "info"].Value;


                var ntlCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlCon);

                dal = new DataLayer();
                dal.Connect(ntlCon);
                SAMPLE sample = null;
                 
                try
                {
                    // Debugger.Launch();
                    long.TryParse(rs.Fields[tableName + "_ID"].Value.ToString(), out tableID);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("This program works on SDG only.");
                    return;
                }


                sample = dal.FindBy<SAMPLE>(s => s.SAMPLE_ID == tableID).SingleOrDefault();
                if (sample == null)
                {
                    MessageBox.Show("Can't find the SAMPLE for the id");
                    return;
                }
                CLIENT client = sample.SDG.SDG_USER.CLIENT;



                if (_appSettings.Settings["SAMPLEPrinter"] != null)
                {

                    string defaultPrinter = _appSettings.Settings["SAMPLEPrinter"].Value;

                    string goodIp = defaultPrinter;


                    string fn = client.CLIENT_USER.U_FIRST_NAME + "  " + client.CLIENT_USER.U_LAST_NAME;

                    string snext = sample.NAME.Substring(10, sample.NAME.Length - 10);
                    string pathoNbr = sample.SDG.SDG_USER.U_PATHOLAB_NUMBER + snext;
                    Print(pathoNbr, sample.NAME, fn,client.NAME );
                }
                else
                    MessageBox.Show("לא הוגדרה מדפסת עבור תחנה זו.");


            }
            catch (Exception ex)
            {
                MessageBox.Show("נכשלה הדפסת מדבקה");
                Logger.WriteLogFile(ex);
            }
            finally
            {
                if (dal != null)
                {
                    dal.Close();
                    dal = null;

                }
            }
        }

        private void Print(string PathoName, string sampleName, string clientFullName, string clientID)
        {

            string ipAddress = _appSettings.Settings["SAMPLEPrinter"].Value;

            string ZPLString;
            try
            {
                ZPLString = _appSettings.Settings["SAMPLEZPLString"].Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't find the app config file with ZPLString entry");
                return;
            }
            PrintOperationCls printOperation = new PrintOperationCls();

            // Reverse hebrew
            clientFullName = printOperation.ManipulateHebrew(clientFullName);

            ZPLString = string.Format(ZPLString, PathoName, sampleName, clientFullName, clientID);
            try
            {
                RawPrinterHelper.SendStringToPrinter(ipAddress, ZPLString);
                return;

            }
            catch (Exception ex)
            {

                RawPrinterHelper.SendStringToPrinter(ipAddress, ZPLString);


            }
        }
    }
}
