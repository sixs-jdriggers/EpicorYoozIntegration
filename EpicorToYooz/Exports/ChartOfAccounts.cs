using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using EpicorRestAPI;
using EpicorToYooz.Properties;
using Newtonsoft.Json.Linq;
using NLog;

namespace EpicorToYooz.Exports {
    internal static class ChartOfAccounts {
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

                Logger.Debug($"Calling BAQ {Settings.Default.BAQ_ChartOfAccounts}...");
                var parameters = new Dictionary<string, string> {{"LastChange", Settings.Default.LastExecution.ToString()}};
                var jsonData   = EpicorRest.GetBAQResultJSON(Settings.Default.BAQ_ChartOfAccounts, parameters);

                var data = JObject.Parse(jsonData)["value"].ToObject<DataTable>();

                Logger.Debug("Converting BAQ results to classes for export...");
                var accounts = data.Rows.Cast<DataRow>()
                                   .Select(row => new Account {
                                       AccountNumber = row[BAQColumns.Calculated_Account.ToString()].ToString(),
                                       Description =
                                           row[BAQColumns.Calculated_Description.ToString()].ToString(),
                                       Classification =
                                           row[BAQColumns.Calculated_Classification.ToString()].ToString(),
                                       Action = row[BAQColumns.Calculated_Action.ToString()].ToString()
                                   })
                                   .ToList();

                Logger.Info("Writing Chart Of Accounts CSV file...");
                using (var writer = new StreamWriter($"{Settings.Default.TempDirectory}\\{Settings.Default.FileName_ChartOfAccounts}"))
                using (var csv = new CsvWriter(writer)) {
                    // Tab delimited field, no column headers
                    csv.Configuration.HasHeaderRecord = false;
                    csv.Configuration.Delimiter       = "\t";
                    csv.Configuration.ShouldQuote     = (field, context) => field.Contains("\t");
                    csv.Configuration.RegisterClassMap<AccountMap>();
                    // Write Special Yooz header data
                    csv.WriteField(Settings.Default.Export_ChartOfAccounts_Version, false);
                    csv.NextRecord();
                    csv.WriteField(Settings.Default.Export_ChartOfAccounts_Header, false);
                    csv.NextRecord();
                    // Export actual records.
                    csv.WriteRecords(accounts);
                }

                Logger.Info("Finished exporting Chart Of Accounts");
            } catch (Exception ex) {
                Logger.Error("Failed to export Chart of Accounts!");
                Logger.Error($"Error: {ex.Message}");
                Logger.Error($"Error: {ex.StackTrace}");
                throw;
            }
        }

        private class Account {
            public string AccountNumber  { get; set; }
            public string Description    { get; set; }
            public string Classification { get; set; }
            public string Action         { get; set; }
        }

        /// <summary>
        ///     This determines the column names and order.
        /// </summary>
        private class AccountMap : ClassMap<Account> {
            public AccountMap() {
                Map(m => m.AccountNumber).Index(0).Name("GL Account Number");
                Map(m => m.Description).Index(1).Name("GL Account Label");
                Map(m => m.Classification).Index(2).Name("Classification");
                Map(m => m.Action).Index(3).Name("Action");
            }
        }

        private enum BAQColumns {
            Calculated_Account,
            Calculated_Description,
            Calculated_Classification,
            Calculated_Action
        }
    }
}