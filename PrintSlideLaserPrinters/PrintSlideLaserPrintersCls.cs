using ADODB;
using LSEXT;
using LSSERVICEPROVIDERLib;
using Oracle.ManagedDataAccess.Client;
using Patholab_Common;
using Patholab_DAL_V1;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PrintSlideLaserPrinters
{

    [ComVisible(true)]
    [ProgId("PrintSlideLaserPrinters.PrintSlideLaserPrintersCls")]
    public class PrintSlideLaserPrintersCls : IWorkflowExtension
    {
        private INautilusServiceProvider _serviceProvider;
        private DataLayer _dataLayer;
        private OracleConnection _connection = null;
        private string _lineToTxtFile;

        public bool DEBUG;

        public void Execute(ref LSExtensionParameters parameters)
        {
            try
            {
                long workflowNodeId = parameters["WORKFLOW_NODE_ID"];
                long aliquotId = GetAliquotId(parameters["RECORDS"]);
                if (aliquotId == 0) return;

                InitializeDatabaseConnection(parameters["SERVICE_PROVIDER"]);

                var aliquot = GetAliquot(aliquotId);
                if (aliquot == null)
                {
                    MessageBox.Show("Can't find the aliquot for the ID.");
                    return;
                }

                PrepareAndWriteLog(aliquot, workflowNodeId);
            }
            catch (Exception ex)
            {
                MessageBox.Show("נכשלה הדפסת סליידים.");
                Logger.WriteLogFile(ex);
            }
            finally
            {
                _dataLayer?.Close();
            }
        }

        private long GetAliquotId(Recordset records)
        {
            try
            {
                records.MoveLast();
                return long.TryParse(records.Fields["ALIQUOT_ID"].Value.ToString(), out long id) ? id : 0;
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
                MessageBox.Show("This program works on ALIQUOT only.");
                return 0;
            }
        }

        private void InitializeDatabaseConnection(INautilusServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
            var ntlCon = Utils.GetNtlsCon(serviceProvider);
            Utils.CreateConstring(ntlCon);
            _dataLayer = new DataLayer();
            _dataLayer.Connect(ntlCon);
            _connection = GetConnection(ntlCon);
        }

        private ALIQUOT GetAliquot(long aliquotId)
        {
            return _dataLayer.FindBy<ALIQUOT>(x => x.ALIQUOT_ID == aliquotId && x.ALIQUOT_USER.U_GLASS_TYPE == "S").FirstOrDefault();
        }

        private void PrepareAndWriteLog(ALIQUOT aliquot, long workflowNodeId)
        {
            var casseteName = GetCasseteName(aliquot);
            var logEntry = CreateLogEntry(aliquot, casseteName);
            var logFilePath = GetPrinterPath(workflowNodeId);

            if (!Directory.Exists(logFilePath))
            {
                MessageBox.Show($"{logFilePath} doesn't exist.");                
                return;
            }

            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssFFF");
            using (StreamWriter writer = File.AppendText(Path.Combine(logFilePath, $"{timestamp}.txt")))
            {
                writer.WriteLine(logEntry);
            }
        }

        private string GetCasseteName(ALIQUOT aliquot)
        {
            return _dataLayer.FindBy<ALIQUOT>(x => x.NAME == aliquot.NAME.Substring(0, aliquot.NAME.Length - 2)).FirstOrDefault()?.NAME;
        }

        private string CreateLogEntry(ALIQUOT aliquot, string casseteName)
        {
            string color = aliquot.ALIQUOT_USER.U_COLOR_TYPE;
            string patholabNumber = aliquot.SAMPLE.SDG.SDG_USER.U_PATHOLAB_NUMBER;
            long? laborant = _dataLayer.FindBy<U_ALL_BARCODE_SCAN_USER>(x => x.U_ENTITY_NAME == casseteName)?.FirstOrDefault()?.U_LABORANT;

            return $"{patholabNumber},{GetSlideExtension(aliquot.NAME)},{color},{laborant}";
        }

        private string GetSlideExtension(string patholabNbr)
        {
            int dotIndex = patholabNbr.IndexOf('.');
            return dotIndex != -1 && dotIndex < patholabNbr.Length - 1 ? patholabNbr.Substring(dotIndex + 1) : string.Empty;
        }

        private string GetPrinterPath(long workflowNodeId)
        {
            string printerEventName = GetParentNodeName(workflowNodeId);
            if (string.IsNullOrEmpty(printerEventName)) return null;

            var entry = _dataLayer.FindBy<PHRASE_ENTRY>(
                pe => pe.PHRASE_HEADER.NAME.Equals("Laser printer") &&
                       pe.PHRASE_INFO.Equals(printerEventName, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

            if (entry != null) return entry.PHRASE_DESCRIPTION;

            MessageBox.Show($"Phrase: Vega Printer. Can't find phrase entry for event: {printerEventName}");
            return null;
        }

        private string GetParentNodeName(long workflowNodeId)
        {
            using (var cmd = new OracleCommand($"SELECT parent_id FROM lims_sys.workflow_node WHERE workflow_node_id={workflowNodeId}", _connection))
            {
                var parentNodeId = cmd.ExecuteScalar();
                if (parentNodeId != null)
                {
                    cmd.CommandText = $"SELECT LONG_NAME FROM lims_sys.workflow_node WHERE workflow_node_id={parentNodeId}";
                    return Convert.ToString(cmd.ExecuteScalar());
                }
            }
            return string.Empty;
        }

        public OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {
            if (ntlsCon == null) return null;

            string connectionString = $"Data Source={ntlsCon.GetServerDetails()};User ID={ntlsCon.GetUsername()};Password={ntlsCon.GetPassword()};";

            if (string.IsNullOrEmpty(ntlsCon.GetUsername()))
            {
                connectionString = "User Id=/;Data Source=" + ntlsCon.GetServerDetails() + ";Connection Timeout=60;";
            }

            var connection = new OracleConnection(connectionString);
            connection.Open();
            SetUserRole(ntlsCon, connection);
            ConnectToSameSession(ntlsCon, connection);

            return connection;
        }

        private void SetUserRole(INautilusDBConnection ntlsCon, OracleConnection connection)
        {
            string limsUserPassword = ntlsCon.GetLimsUserPwd();
            string roleCommand = string.IsNullOrEmpty(limsUserPassword)
                ? "SET ROLE lims_user"
                : $"SET ROLE lims_user IDENTIFIED BY {limsUserPassword}";

            using (var command = new OracleCommand(roleCommand, connection))
            {
                command.ExecuteNonQuery();
            }
        }

        private void ConnectToSameSession(INautilusDBConnection ntlsCon, OracleConnection connection)
        {
            double sessionId = ntlsCon.GetSessionId();
            using (var command = new OracleCommand($"CALL lims.lims_env.connect_same_session({sessionId})", connection))
            {
                command.ExecuteNonQuery();
            }
        }
    }
}
