using System;
using System.Data;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using EpicorRestAPI;
using EpicorToYooz.Properties;
using NLog;

namespace EpicorToYooz.Exports {
    internal class POs {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void Export() {
            try {
                Logger.Debug("Setting up Epicor REST library...");
                EpicorRest.AppPoolHost     = Settings.Default.Epicor_Server;
                EpicorRest.AppPoolInstance = Settings.Default.Epicor_Instance;
                EpicorRest.UserName        = Settings.Default.Epicor_User;
                EpicorRest.Password        = Settings.Default.Epicor_Pass;

                Logger.Debug($"Calling BAQ {Settings.Default.BAQ_POs}...");
                var data = EpicorRest.GetBAQResults(Settings.Default.BAQ_POs, null);

                Logger.Debug("Converting BAQ results to classes for export...");
                var orders = data.Rows.Cast<DataRow>().Select(row => new Account {Company = row[BAQColumns.GLAccount_Company.ToString()].ToString()}).ToList();

                Logger.Info("Writing Purchase Order CSV file...");
                using (var writer = new StreamWriter($"{Settings.Default.TempDirectory}\\{Settings.Default.FileName_POs}"))
                using (var csv = new CsvWriter(writer)) {
                    csv.Configuration.RegisterClassMap<POMap>();
                    csv.WriteRecords(orders);
                }

                Logger.Info("Finished exporting Purchase Orders");
            } catch (Exception ex) {
                Logger.Error("Failed to export Purchase Orders!");
                Logger.Error($"Error: {ex.Message}");
                throw;
            }
        }

        internal class PO {
            public string Company           { get; set; }
            public string GL_Account_Number { get; set; }
            public string GL_Account_Label  { get; set; }
            public string Classification    { get; set; }
            public string Action            { get; set; }
        }

        /// <summary>
        ///     This determines the column names and order.
        /// </summary>
        private class POMap : ClassMap<PO> {
            public POMap() {
                Map(m => m.Company).Index(0).Name("Company");
                Map(m => m.GL_Account_Number).Index(1).Name("GL PO Number");
                Map(m => m.GL_Account_Label).Index(2).Name("GL PO Label");
                Map(m => m.Classification).Index(3).Name("Classification");
                Map(m => m.Action).Index(4).Name("Action");
            }
        }

        private enum BAQColumns {
            GLAccount_Company,
            GLAccount_GLAccount,
            GLAccount_AccountDesc,
            Calculated_Classification,
            Calculated_Action
        }
    }
}