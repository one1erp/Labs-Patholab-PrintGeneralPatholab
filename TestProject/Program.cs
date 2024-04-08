using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32.SafeHandles;
using PrintGeneralPatholab;


namespace TestProject
{
    internal class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
        uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);
        private static int _port = 9100;
        private static AppSettingsSection _appSettings;
        private static void Main(string[] args)
        {
            string assemblyPath = @"C:\Users\roye\Desktop\Patholab Projects\PrintGeneralPatholab\PrintReceivedLabels\bin\Debug\PrintReceivedLabels.dll";
                ExeConfigurationFileMap map = new ExeConfigurationFileMap ();

            map.ExeConfigFilename = assemblyPath + ".config";
            Configuration cfg = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
            _appSettings = cfg.AppSettings;
            Print("123456/16", "123456/16", "נתן משה", "123456789");
        }

        private static void Print(string PathoName, string sdgName, string clientFullName, string clientName)
        {

            string ipAddress = _appSettings.Settings["defaultPAPPrinter"].Value;

            string ZPLString;
            try
            {
                ZPLString = _appSettings.Settings["ZPLStringPAP"].Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't find the app config file with ZPLString entry");
                return;
            }
            PrintOperationCls printOperation = new PrintOperationCls();

            int lettersInaRow = 16;
            if (clientFullName.Length > lettersInaRow)
                clientFullName = clientFullName.Substring(0, lettersInaRow);

            // Reverse hebrew
            clientFullName = printOperation.ManipulateHebrew(clientFullName);

            
         //   ZPLString = string.Format(ZPLString,
                //CenterByAddingBlanks(""+ PathoName, lettersInaRow),
                //sdgName,
                //CenterByAddingBlanks(clientFullName,lettersInaRow),
                //CenterByAddingBlanks(""+clientName,lettersInaRow)
                //);
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


        private static string CenterByAddingBlanks(string clinicName,int lettersInARow)
        {
            string blanks = "                                      ";
            if (clinicName.Length < lettersInARow)
            {
                blanks = blanks.Substring(0, (int)Math.Round((lettersInARow - clinicName.Length) / 2.0));
                clinicName = blanks + clinicName;
            }
            return clinicName;
        }
    }
}

