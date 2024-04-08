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
    [ProgId("PrintReceivedLabels.PrintReceivedLabelsCls")]
    public class PrintReceivedLabelsCls : IWorkflowExtension
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
        private long _operatorId;
        private bool debug;



        public void Execute(ref LSExtensionParameters Parameters)

        {
            try
            {
                #region param
                string role = Parameters["ROLE_NAME"];


                debug = (role.ToUpper() == "DEBUG");

                if (debug) Debugger.Launch();

                string tableName = Parameters["TABLE_NAME"];
                sp = Parameters["SERVICE_PROVIDER"];
                var rs = Parameters["RECORDS"];
                _operatorId =(long) Utils.GetNautilusUser(sp).GetOperatorId();
                rs.MoveLast();
                long tableID;
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
                if (tableName != "SDG")
                {
                    MessageBox.Show("This program works on SDG only.");
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
                SDG sdg = null;


                sdg = dal.FindBy<SDG>(c => c.SDG_ID == tableID).Include("SDG_USER").SingleOrDefault();





                if (sdg == null)
                {
                    MessageBox.Show("Can't find the SDG for the id");
                    return;
                }
                SDG_USER sdgUser = sdg.SDG_USER;

                CLIENT client = sdgUser.CLIENT;

                if (_appSettings.Settings[tableName + "Printer"] != null)
                {

                    string defaultPrinter = _appSettings.Settings[tableName+"Printer"].Value;

                    string goodIp = defaultPrinter;
                    string createdOnString = "";

                    if (sdg.CREATED_ON !=null)
                    { 
                        createdOnString = ((DateTime)sdg.CREATED_ON).ToString("dd/MM/yy");
                    }
                    
                   
                    string fn = client.CLIENT_USER.U_FIRST_NAME + "  " + client.CLIENT_USER.U_LAST_NAME;

                    if (sdg.NAME.Substring(0, 1) != "P")
                    {

                        Print(sdgUser.U_PATHOLAB_NUMBER, sdg.NAME, client.NAME, fn, sdg.SAMPLEs.Count, createdOnString);
                    }
                    else
                    {
                        PrintPAP(sdgUser.U_PATHOLAB_NUMBER, sdg.NAME, client.NAME, fn, sdg.SAMPLEs.Count,createdOnString, sdg);
                    }
                   
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

        private void Print(string PathoName, string sdgName, string clientName, string clientFullName, int samplesSum, string receivedOn)
        {
            //how did it work with defalut printer?
            string ipAddress = _appSettings.Settings["SDGPrinter"].Value;

            string ZPLString;
            try
            {
                ZPLString = _appSettings.Settings["SDGZPLString"].Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't find the app config file with SDGZPLString entry");
                return;
            }
            PrintOperationCls printOperation = new PrintOperationCls();

            // Reverse hebrew
            clientFullName = printOperation.ManipulateHebrew(clientFullName);
                                        //origin , 0          1           2           3           4            5
            ZPLString = string.Format(ZPLString, PathoName, sdgName, clientFullName,clientName,  samplesSum, receivedOn);
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

        private void PrintPAP(string PathoName, string sdgName, string clientName, string clientFullName, int samplesSum, string receivedOn, SDG sdg)
        {

            string ipAddress = _appSettings.Settings["defaultPAPPrinter"].Value;

            string ZPLStringPAP ;
            string ZPLStringPAPNoAliquot;
            try   
            {
                ZPLStringPAP = _appSettings.Settings["ZPLStringPAP"].Value;
                ZPLStringPAPNoAliquot = _appSettings.Settings["ZPLStringPAPNoAliquot"].Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't find the app config file with ZPLStringPAP or ZPLStringPAPNoAliquot entries");
                return;
            }
            PrintOperationCls printOperation = new PrintOperationCls();

            // Reverse hebrew
            clientFullName = printOperation.ManipulateHebrew(clientFullName);
            ALIQUOT[] slides;
            try
            {
                //.where(a=>a.ALIQUOT_USER.U_GLASS_TYPE="S")
                 slides =
                    sdg.SAMPLEs.OrderBy(x => x.NAME).Where(d=>d.STATUS!="X").FirstOrDefault()
                    .ALIQUOTs.Where(a=>a.ALIQUOT_USER.U_GLASS_TYPE=="S" && a.STATUS!="X")
                    .OrderBy(a => a.NAME).ToArray();
            
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
                slides = null;
            }

            string ZPLString;
            if(slides==null || slides.Count()==0)
            {
                ZPLString = string.Format(ZPLStringPAPNoAliquot, PathoName, sdgName);
                SendToPrinter(ipAddress, ZPLString);
            }
            else
                foreach (ALIQUOT slide in slides)
                {
                    {
                        // here i print the lable for the sdg and slide. 
                        // in case this a pap LBC "duck" (thin prep) i will print the sample data in the lable
                        string barcode = slide.NAME;
                        try
                        {
                            if (slide.NAME.Split('.')[1] != "0") //pap LBC
                            {
                                barcode = slide.SAMPLE.NAME;
                                if (slide.NAME.Split('.')[3] != "1") //pap LBC - print for the first slide
                                {
                                    continue;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteLogFile(ex);
                            continue;
                        }
                        string slideNameAterFirstDot = barcode.Substring(barcode.IndexOf(('.')));

                        string SlideorSampleHeader = PathoName + slideNameAterFirstDot;
                        string colorType = (slide.ALIQUOT_USER.U_COLOR_TYPE ?? "");
                        if (colorType.Length > 10)
                            colorType = colorType.Substring(0, 10);
                        colorType = printOperation.ManipulateHebrew(colorType);
                        string pathoNumber =
                            dal.FindBy<OPERATOR_USER>(o => o.OPERATOR_ID == _operatorId)
                               .SingleOrDefault()
                               .U_PATHOLAB_WORKER_NBR ??
                            "00";
                        //origin , 0          1           2           3           4     5
                        ZPLString = string.Format(ZPLStringPAP, PathoName, sdgName, barcode, SlideorSampleHeader,
                                                  colorType, pathoNumber);
                        SendToPrinter(ipAddress, ZPLString);
                    }
                }
           
        }

        private static void SendToPrinter(string ipAddress, string ZPLString)
        {
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
