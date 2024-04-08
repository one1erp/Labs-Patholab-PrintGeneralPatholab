using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using LSEXT;
using Patholab_Common;
using Patholab_DAL_V1;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using PrintGeneralPatholab;
using VentanaHL7;
namespace PrintSlideNew
{

    [ComVisible(true)]
    [ProgId("PrintSlideNew.PrintSlideNewCls")]
    public class PrintSlideNewCls : IWorkflowExtension
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
        private VentanaHL7X ventana;
        private VentanaOrderX newOrder;
        private AutoResetEvent autoResetEvent;
        private string ventanaip = "192.168.0.38";
        private int port = 58000;
        private bool _firstConnection = true;
        private ManualResetEvent manualResetEvent = new ManualResetEvent(false);
        private string ventanaBarcode;
        private int TimeToWaitForErrorMessageInMs = 1000;
        private bool debug;

        private ALIQUOT slide;

        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {

                #region param

                string tableName = Parameters["TABLE_NAME"];
                sp = Parameters["SERVICE_PROVIDER"];
                var rs = Parameters["RECORDS"];
                rs.MoveLast();
                long tableID;
                try
                {
                    // Debugger.Launch();
                    long.TryParse(rs.Fields["ALIQUOT_ID"].Value.ToString(), out tableID);
                }
                catch (Exception ex)
                {
                    TimedMessageBox("This program works on aliqout only.");
                    return;
                }

                var workstationId = Parameters["WORKSTATION_ID"];


                #endregion

                if (Parameters["ROLE_NAME"].ToString().ToUpper() == "DEBUG")
                {
                    debug = true;
                    Debugger.Launch();
                }


                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                ExeConfigurationFileMap map = new ExeConfigurationFileMap();
                map.ExeConfigFilename = assemblyPath + ".config";
                Configuration cfg = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                _appSettings = cfg.AppSettings;


                var ntlCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlCon);

                dal = new DataLayer();
                dal.Connect(ntlCon);



                slide = dal.FindBy<ALIQUOT>(a => a.ALIQUOT_ID == tableID).SingleOrDefault();

                if (slide.ALIQUOT_USER.U_GLASS_TYPE != "S")
                {
                    TimedMessageBox("Can't Print, Glass Type is not 'S'");
                    return;
                }

                string printerName = null;
                string workerNumber = "";
                U_ALL_BARCODE_SCAN_USER barcodeScanUser = null;


                string defaultPrinter = _appSettings.Settings["defaultPrinter"].Value;

                if (slide.ALIQ_FORMULATION_PARENT != null && slide.ALIQ_FORMULATION_PARENT.Count > 0)
                {
                    ALIQUOT block = slide.ALIQ_FORMULATION_PARENT.SingleOrDefault().CHILDREN;
                    if (block != null)
                    {
                        //look for block
                        barcodeScanUser = dal.FindBy<U_ALL_BARCODE_SCAN_USER>
                            (b => b.U_ENTITY_NAME == block.NAME
                             && b.U_CREATED_ON > DateTime.Today)
                              .OrderByDescending(b => b.U_CREATED_ON).FirstOrDefault();
                    }
                }
                else
                {
                    //if block not found, look for sample (PAP or CYT)
                    barcodeScanUser = dal.FindBy<U_ALL_BARCODE_SCAN_USER>
                        (b => b.U_ENTITY_NAME == slide.SAMPLE.NAME
                         && b.U_CREATED_ON > DateTime.Today)
                         .OrderByDescending(b => b.U_CREATED_ON).FirstOrDefault();
                }


                if (barcodeScanUser != null)
                {
                    printerName = barcodeScanUser.U_PRINTER;
                    workerNumber = barcodeScanUser.OPERATOR.OPERATOR_USER.U_PATHOLAB_WORKER_NBR;

                }
                else
                {
                    if (printerName == null)
                    {
                        printerName = defaultPrinter;
                    }
                    else
                    {
                        TimedMessageBox("לא הוגדרה מדפסת עבור תחנה זו.");
                        return;
                    }
                    //    TimedMessageBox("Can't find the Scan Record for the block");

                }
                string patholabNumber = slide.SAMPLE.SDG.SDG_USER.U_PATHOLAB_NUMBER;
                string slideNameAterFirstDot = slide.NAME.Substring(slide.NAME.IndexOf(('.')));
                string SlideHeader = patholabNumber + slideNameAterFirstDot;
                string colorType = (slide.ALIQUOT_USER.U_COLOR_TYPE ?? "");

                //   Debugger.Launch ( );
                var exreq = (from item in dal.FindBy<U_EXTRA_REQUEST_DATA_USER>(x => x.U_SLIDE_NAME == slide.NAME
                               && x.U_EXTRA_REQUEST.U_EXTRA_REQUEST_USER.U_SDG_ID == slide.SAMPLE.SDG.SDG_ID && x.U_ENTITY_TYPE == "Block"
                               )

                             select item.U_EXTRA_REQUEST.U_EXTRA_REQUEST_USER.U_CREATED_BY).FirstOrDefault();

                string opName = "";
                if (exreq != null)
                {
                    var oPERATOR = dal.FindBy<OPERATOR>(x => x.OPERATOR_ID == exreq.Value).FirstOrDefault();
                    if (oPERATOR != null)
                    {
                        opName = oPERATOR.FULL_NAME;//.name; Netanel asked 21-03-22
                    }

                }

                //     select item.U_EXTRA_REQUEST.U_EXTRA_REQUEST_USER.U_CREATED_BY).FirstOrDefault();
                //check for ventana colors
                //Debugger.Launch();
                U_PARTS part = dal.FindBy<U_PARTS>(p => p.U_PARTS_USER.U_STAIN == colorType).FirstOrDefault();
               
                var protocolName = part!=null?part.U_PARTS_USER.U_PROTOCOL_NUMBER.ToString():"";

                if (part != null && part.U_PARTS_USER.U_PART_TYPE == "I" && part.DESCRIPTION.Contains("^STAIN"))
                {
                    if (!PrintVentana(SlideHeader, slide.NAME, workerNumber, colorType, printerName))
                        Print(SlideHeader, slide.NAME, workerNumber, colorType, opName, printerName, protocolName);

                }
                else
                {
                    Print(SlideHeader, slide.NAME, workerNumber, colorType, opName, printerName, protocolName);
                }
                if (barcodeScanUser != null)
                    slide.ALIQUOT_USER.U_LAST_LABORANT = barcodeScanUser.U_LABORANT;
                slide.ALIQUOT_USER.U_PRINTED_ON = dal.GetSysdate();
                dal.SaveChanges();

            }
            catch (Exception ex)
            {
                TimedMessageBox("נכשלה הדפסת מדבקה");
                Logger.WriteLogFile(ex);
            }
            finally
            {
                if (dal != null)
                {
                    dal.Close();
                }
            }
        }

        private void Print(string header, string Barcode, string workerNumber, string colorType, string opName, string printerName, string protocolName)
        {

            string defaultIpAddress = _appSettings.Settings [ "defaultPrinter" ].Value;

            string ZPLString;
            try
            {
                ZPLString = _appSettings.Settings [ "ZPLString" ].Value;
            }
            catch ( Exception ex )
            {
                TimedMessageBox ( "Can't find the app config file with ZPLString entry" );
                return;
            }
            PrintOperationCls printOperation = new PrintOperationCls ( );
            // Reverse hebrew

            //only print the firs 10 letters
            if ( colorType.Length > 13 )
                colorType = colorType.Substring ( 0, 13 );
            colorType = printOperation.ManipulateHebrew ( colorType );
            //  decimal width = (19 - header.Length) * 1 + 19;

            //  string headerwidth = Convert.ToString(Math.Round(width));
            if ( ZPLString.Contains ( "{4}" ) )
            {
                string headerAfterDot = header.Substring ( header.IndexOf ( "." ) );
                header = header.Substring ( 0, header.IndexOf ( "." ) );
                ZPLString = string.Format(ZPLString, header, Barcode, workerNumber, colorType, headerAfterDot, opName, protocolName);//Get operator from extra request
            }
            else
            {
                ZPLString = string.Format(ZPLString, header, Barcode, workerNumber, colorType, opName, protocolName);

            }
            try
            {
                RawPrinterHelper.SendStringToPrinter ( printerName, ZPLString );

                // Open connection

            }
            catch ( Exception ex )
            {
                //   PrintDialog pd  = new PrintDialog();
                //  pd.PrinterSettings = new PrinterSettings();
                //if (DialogResult.OK == pd.ShowDialog())0
                //{

                RawPrinterHelper.SendStringToPrinter ( defaultIpAddress, ZPLString );

                //}
            }
        }
        private bool PrintVentana(string header, string Barcode, string workerNumber, string colorType, string printerName)
        {

            string ZPLStringWithVentana;
            int TimeToWaitForVentanaServiceInMs;
            try
            {
                ventanaip = _appSettings.Settings["ventanaip"].Value;
            }
            catch (Exception ex)
            {
                TimedMessageBox("Can't find the app config file with ventanaip entry");
                return false;
            }
            try
            {
                ZPLStringWithVentana = _appSettings.Settings["ZPLStringWithVentana"].Value;
            }
            catch (Exception ex)
            {
                TimedMessageBox("Can't find the app config file with ZPLStringWithVentana entry");
                return false;
            }
            try
            {
                TimeToWaitForVentanaServiceInMs = int.Parse(_appSettings.Settings["TimeToWaitForVentanaServiceInMs"].Value);
            }
            catch (Exception ex)
            {
                TimedMessageBox("Can't find the app config file with TimeToWaitForVentanaServiceInMs entry");
                return false;
            }
            try
            {
                TimeToWaitForErrorMessageInMs = int.Parse(_appSettings.Settings["TimeToWaitForErrorMessageInMs"].Value);
            }
            catch (Exception ex)
            {
                TimedMessageBox("Can't find the app config file with TimeToWaitForErrorMessageInMs entry");
                return false;
            }
            try
            {


                ThreadStart ventanaRef = new ThreadStart(ventanaConnect);
                Thread ventanaThread = new Thread(ventanaRef);
                ventanaThread.Start();
                if (!manualResetEvent.WaitOne(TimeToWaitForVentanaServiceInMs))
                {
                    ventanaThread.Abort();
                    try
                    {
                        ventana.Close();
                    }
                    catch (Exception e)
                    {
                    }
                    TimedMessageBox("Print Ventana action Faild \r\nBarcode:" + Barcode + "\r\nColor:" + colorType);
                    Logger.WriteLogFile(
                        new Exception("Print Ventana action Faild \r\nBarcode:" + Barcode + "\r\nColor:" + colorType));

                    return false;
                }
            }
            catch (Exception e)
            {
                Logger.WriteLogFile(e);
            }
            finally
            {

                try
                {
                    ventana.Close();
                }
                catch (Exception e)
                {
                }
            }
            string defaultIpAddress = _appSettings.Settings["defaultPrinter"].Value;

            PrintOperationCls printOperation = new PrintOperationCls();
            // Reverse hebrew

            //only print the firs 10 letters
            if (colorType.Length > 10)
                colorType = colorType.Substring(0, 10);
            colorType = printOperation.ManipulateHebrew(colorType);
            decimal width = (19 - header.Length) * 1 + 19;
            string headerwidth = Convert.ToString(Math.Round(width));
            if (string.IsNullOrEmpty(ventanaBarcode))
            {
                TimedMessageBox("Ventana code not found \r\nBarcode:" + Barcode);
            }
            ventanaBarcode = GetBarcodeWithITF(ventanaBarcode);
            slide.EXTERNAL_REFERENCE = ventanaBarcode;
            dal.SaveChanges();
            if (string.IsNullOrEmpty(ventanaBarcode))
            {
                TimedMessageBox("Ventana code not numeric \r\nBarcode:" + Barcode);
            }
            ZPLStringWithVentana = string.Format(ZPLStringWithVentana, header, Barcode, workerNumber, colorType,
                                                 headerwidth, ventanaBarcode);
            try
            {
                RawPrinterHelper.SendStringToPrinter(printerName, ZPLStringWithVentana);
                return true;

            }
            catch (Exception ex)
            {

                RawPrinterHelper.SendStringToPrinter(defaultIpAddress, ZPLStringWithVentana);
                return true;
                //}
            }

        }
        private void TimedMessageBox(string message, int wait)
        {

            Logger.WriteLogFile(new Exception(message));
            if (debug)
            {
                //   Form1 form1 = new Form1(message, wait);
                //   form1.Show();
            }
        }
        private void TimedMessageBox(string message)
        {
            Logger.WriteLogFile(new Exception(message));

            if (debug)
            {
                //   Form1 form1 = new Form1(message, TimeToWaitForErrorMessageInMs);
                //  form1.Show();
            }
        }
        private string GetBarcodeWithITF(string barcode)
        {
            int i;
            if (!int.TryParse(barcode, out i))
            {
                return null;
            }
            i = 0;
            int sum = 0;
            int digit;
            bool odd = true;// firs is odd
            foreach (char cdigit in barcode)
            {
                digit = int.Parse(cdigit.ToString());
                if (odd)
                {
                    sum += digit;
                }
                else
                {
                    sum += digit * 2;
                }
                odd = !odd;
            }
            int checksum = sum % 10;
            checksum = (10 - checksum) % 10;
            return barcode + checksum.ToString();
        }

        private void ventanaConnect()
        {
            ventana = new VentanaHL7X();
            _firstConnection = true;
            ventana.OnConnection += ventana_OnConnection;
            ventana.OnLogEvent += ventana_OnLogEvent;
            ventana.Open(ventanaip, port);
        }
        private void ventana_OnConnection()
        {
            if (_firstConnection)
            {
                try
                {


                    _firstConnection = false;

                    //ventana.QueryProtocols();
                    //22^MELAN A^STAIN
                    //ventana.QueryTemplates();
                    ///todo: move this outside , befor the timer

                    if (slide != null && slide.ALIQUOT_USER.U_COLOR_TYPE != null)
                    {
                        U_PARTS part = dal.FindBy<U_PARTS>(p => p.U_PARTS_USER.U_STAIN == slide.ALIQUOT_USER.U_COLOR_TYPE
                                                                && p.U_PARTS_USER.U_PART_TYPE == "I").FirstOrDefault();
                        if (part != null)
                        {
                            //part.DESCRIPTION is the ventana staining code
                            string ventanaStainCode = part.DESCRIPTION;

                            if (string.IsNullOrEmpty(ventanaStainCode) || !ventanaStainCode.Contains("^STAIN"))
                            {
                                TimedMessageBox("Cannot find the correct stain code fo the Imuno color type: " +
                                                slide.ALIQUOT_USER.U_COLOR_TYPE + "\r\nPrint Canceled!");
                                return;
                            }
                            newOrder = ventana.NewOrder(slide.ALIQUOT_ID.ToString(), ventanaStainCode, "Patho-Lab", true, true,
                                                        true, true);
                            //Case ID
                            newOrder.SetFieldValue("ORC.2", slide.NAME);

                            newOrder.SetFieldValue("ORC.21", "P^Patho Lab");
                            //Requester
                            newOrder.SetFieldValue("PV1.7", "1^LIS^Nautilus^1");

                            newOrder.SetFieldValue("OBR.1", "1");

                            newOrder.PlaceOrder();
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {

                    Logger.WriteLogFile(ex);

                }
            }
        }


        void ventana_OnLogEvent(string AText)
        {
            //result
            //MSH|^~\&|VIP|Pathology Lab|LIS|Pathology Lab|20160709233647448||ORL^O22|VMSG2|P|2.4|
            //MSA|AA|MSG1|
            //PV1|||||||1^LIS^Nautilus^1|
            //SAC|
            //ORC|OK|TEST69/16|00073||ID|L~B~E~I|||||||||||||||P^Patho Lab|
            //OBR|1|1269|00073|22^MELAN A^STAIN||||||||||||||Patho-Lab|
            if (newOrder == null) return;
            if (AText.Contains("|ORL^O22|"))
            {
                //this is a result
                //Take the OBR part
                string obr = AText.Split(new string[] { "\rOBR|" }, options: StringSplitOptions.None)[1];
                //split it. the nautilus I.D is 1, the internal ID is 2
                string[] obrArray = obr.Split(new char[] { '|' }, StringSplitOptions.None);
                if (newOrder.PlacerSlideID == obrArray[1])
                {
                    //     ventana.QueryOrder(obrArray[2],true,true);
                    //UpdateDbAndPrint();

                    ventanaBarcode = obrArray[2];
                    manualResetEvent.Set();

                }

            }
        }

    }
}
