using LSEXT;
using LSSERVICEPROVIDERLib;
using Microsoft.Win32.SafeHandles;
using Patholab_Common;
using Patholab_DAL_V1;
using PrintGeneralPatholab;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PrintContainer
{

    [ComVisible(true)]
    [ProgId("PrintContainer.PrintContainerClsEntEx")]
    public class PrintContainerClsEntEx : IEntityExtension
    {
        private INautilusServiceProvider sp;
        private const string Type = "3";
        private int _port = 9100;

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
                                                        uint dwShareMode, IntPtr lpSecurityAttributes,
                                                        FileMode dwCreationDisposition,
                                                        uint dwFlagsAndAttributes, IntPtr hTemplateFile);

        private DataLayer dal;
        private AppSettingsSection _appSettings;




        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {


      


                #region Init
                Logger.WriteQueries("start class");

                //   string tableName = Parameters["TABLE_NAME"];
                sp = Parameters["SERVICE_PROVIDER"];
                var records = Parameters["RECORDS"];
                var ntlCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlCon);

                dal = new DataLayer();
                dal.Connect(ntlCon);
                #endregion

                #region Config
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                ExeConfigurationFileMap map = new ExeConfigurationFileMap();
                map.ExeConfigFilename = assemblyPath + ".config";
                Configuration cfg = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                _appSettings = cfg.AppSettings;
                string aaaa = _appSettings.Settings["info"].Value;
                #endregion
                while (!records.EOF)
                {
                    string a = records.Fields["U_CONTAINER_ID"].Value.ToString();
                    long id = long.Parse(a);
                    U_CONTAINER container = dal.FindBy<U_CONTAINER>(c => c.U_CONTAINER_ID == id).SingleOrDefault();
                    U_CLINIC clinic = container.U_CONTAINER_USER.U_CLINIC1;
                    MessageBox.Show(a);
                    if (_appSettings.Settings["defaultPrinter"] != null)
                    {

                        string defaultPrinter = _appSettings.Settings["defaultPrinter"].Value;
                        string goodIp = defaultPrinter;

                        string clinicName = "";
                        try
                        {
                            clinicName = container.U_CONTAINER_USER.U_CLINIC1.NAME;
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteLogFile(ex);
                        }

                        string name = container.NAME;
                        string recievedOn = container.U_CONTAINER_USER.U_RECEIVED_ON.ToString();


                        Print(name, name, clinicName,
                              (container.U_CONTAINER_USER.U_NUMBER_OF_ORDERS ?? 0).ToString(),
                              (container.U_CONTAINER_USER.U_NUMBER_OF_SAMPLES ?? 0).ToString(),
                              goodIp);
                 

                    }
                    else
                        MessageBox.Show("לא הוגדרה מדפסת עבור תחנה זו.");
                    //here
                    records.MoveNext();
                }





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

        private void Print(string header, string Barcode, string clinicName, string jarsQuantity, string hafnayot,
                                  string ip)
        {

            string ipAddress = _appSettings.Settings["defaultPrinter"].Value;

            string ZPLString;
            try
            {
                ZPLString = _appSettings.Settings["ZPLStringContainer"].Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't find the app config file with ZPLString entry");
                Logger.WriteLogFile(ex);
                return;
            }
            PrintOperationCls printOperation = new PrintOperationCls();
            // Reverse hebrew

            //only print the firs 15 letters
            if (clinicName.Length > 15)
                clinicName = clinicName.Substring(0, 15);
            clinicName = printOperation.ManipulateHebrew(clinicName);

            //add banks to clinic name and header
            clinicName = CenterByAddingBlanks(clinicName, 15);
            header = CenterByAddingBlanks(header, 15);
            // header = "  " + header;

            ZPLString = string.Format(ZPLString, header, Barcode, clinicName, jarsQuantity, hafnayot);
            try
            {
                RawPrinterHelper.SendStringToPrinter(ipAddress, ZPLString);
                return;
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
                RawPrinterHelper.SendStringToPrinter(ipAddress, ZPLString);


            }
        }

        private string CenterByAddingBlanks(string clinicName, int lettersInARow)
        {
            string blanks = "                                      ";
            if (clinicName.Length < lettersInARow)
            {
                blanks = blanks.Substring(0, (int)Math.Round((lettersInARow - clinicName.Length) / 2.0));
                clinicName = blanks + clinicName;
            }
            return clinicName;
        }

        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exEnabled;
        }
    }
}
