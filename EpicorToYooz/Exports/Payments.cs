using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml;
using CsvHelper;
using CsvHelper.Configuration;
using EpicorRestAPI;
using EpicorToYooz.Properties;
using Newtonsoft.Json.Linq;
using NLog;

namespace EpicorToYooz.Exports {
    internal static class Payments {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void Export() {
            try {
                Logger.Debug("Setting up Epicor REST library...");
                EpicorRest.AppPoolHost      = Settings.Default.Epicor_Server;
                EpicorRest.AppPoolInstance  = Settings.Default.Epicor_Instance;
                EpicorRest.UserName         = Settings.Default.Epicor_User;
                EpicorRest.Password         = Settings.Default.Epicor_Pass;
                EpicorRest.CallSettings     = new CallSettings(Settings.Default.Epicor_Company, string.Empty, string.Empty, string.Empty);
                EpicorRest.IgnoreCertErrors = true;

                Logger.Debug($"Calling BAQ {Settings.Default.BAQ_Payments}...");
                var parameters = new Dictionary<string, string> {{"LastChange", Settings.Default.LastExecution.ToString()}};
                var jsonData   = EpicorRest.GetBAQResultJSON(Settings.Default.BAQ_Payments, parameters);

                var data = JObject.Parse(jsonData)["value"].ToObject<DataTable>();

                Logger.Debug("Converting BAQ results to classes for export...");
                var payments = data.Rows.Cast<DataRow>()
                                   .Select(row => new Payment {
                                       DocumentID = row[BAQColumns.Calculated_DocumentID.ToString()].ToString(),
                                       Reference  = row[BAQColumns.APTran_CheckNum.ToString()].ToString(),
                                       Date       = row[BAQColumns.APTran_TranDate.ToString()].ToString(),
                                       Amount     = row[BAQColumns.Calculated_Amount.ToString()].ToString()
                                   })
                                   .ToList();

                if (payments.Count==0) {
                    Logger.Info("No payments found. Skipping file generation.");
                    return;
                }

                Logger.Info("Writing Payments XML file...");
                using (var writer = XmlWriter.Create($"{Settings.Default.TempDirectory}\\{Settings.Default.FileName_Payments}")) {
                    writer.WriteStartDocument();
                    // Message line
                    writer.WriteStartElement("message");
                    writer.WriteAttributeString("xmlns", "xsi", "", "http://www.w3.org/2001/XMLSchema-instance");
                    writer.WriteAttributeString("xmlns", "xsd", "", "http://www.w3.org/2001/XMLSchema");
                    writer.WriteAttributeString("type",      "GENERIC_FIELDS");
                    
                    // Records
                    foreach (var payment in payments) {
                        // Start Document
                        writer.WriteStartElement("document");
                        writer.WriteAttributeString("id", payment.DocumentID);
                        
                        // Fields
                        // Reference
                        writer.WriteStartElement("field");
                        writer.WriteAttributeString("name","PAYMENT_REFERENCE");
                        writer.WriteAttributeString("value",payment.Reference);
                        writer.WriteEndElement();
                        // Date
                        writer.WriteStartElement("field");
                        writer.WriteAttributeString("name",  "PAYMENT_DATE");
                        var date = DateTime.Parse(payment.Date);
                        writer.WriteAttributeString("value", date.ToString("yyyyMMdd"));
                        writer.WriteEndElement();
                        // Amount
                        writer.WriteStartElement("field");
                        writer.WriteAttributeString("name",  "PAYMENT_AMOUNT");
                        writer.WriteAttributeString("value", decimal.Parse(payment.Amount).ToString("F2"));
                        writer.WriteEndElement();

                        // End Document
                        writer.WriteEndElement();
                    }

                    // Close up the file
                    writer.WriteEndDocument();
                }

                Logger.Info("Finished exporting Payments");
            } catch (Exception ex) {
                Logger.Error("Failed to export Payments!");
                Logger.Error($"Error: {ex.Message}");
                Logger.Error($"Error: {ex.StackTrace}");
                throw;
            }
        }

        private class Payment {
            public string DocumentID { get; set; }
            public string Date       { get; set; }
            public string Reference  { get; set; }
            public string Amount     { get; set; }
        }

        private enum BAQColumns {
            Calculated_DocumentID,
            APTran_CheckNum,
            APTran_TranDate,
            Calculated_Amount
        }
    }
}