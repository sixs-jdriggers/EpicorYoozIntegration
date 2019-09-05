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
    internal static class ChartOfAccounts {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void Export() {
            try {
                Logger.Debug("Setting up Epicor REST library...");
                EpicorRest.AppPoolHost     = Settings.Default.Epicor_Server;
                EpicorRest.AppPoolInstance = Settings.Default.Epicor_Instance;
                EpicorRest.UserName        = Settings.Default.Epicor_User;
                EpicorRest.Password        = Settings.Default.Epicor_Pass;

                Logger.Debug($"Calling BAQ {Settings.Default.BAQ_ChartOfAccounts}...");
                var data = EpicorRest.GetBAQResults(Settings.Default.BAQ_ChartOfAccounts, null);

                Logger.Debug("Converting BAQ results to classes for export...");
                var accounts = data.Rows.Cast<DataRow>()
                                   .Select(row => new Account {
                                       Company           = row[BAQColumns.GLAccount_Company.ToString()].ToString(),
                                       GL_Account_Number = row[BAQColumns.GLAccount_GLAccount.ToString()].ToString(),
                                       GL_Account_Label  = row[BAQColumns.GLAccount_AccountDesc.ToString()].ToString(),
                                       Classification =
                                           row[BAQColumns.Calculated_Classification.ToString()].ToString(),
                                       Action = row[BAQColumns.Calculated_Action.ToString()].ToString()
                                   })
                                   .ToList();

                Logger.Info("Writing Chart Of Accounts CSV file...");
                using (var writer = new StreamWriter($"{Settings.Default.TempDirectory}\\{Settings.Default.FileName_ChartOfAccounts}"))
                using (var csv = new CsvWriter(writer)) {
                    csv.Configuration.RegisterClassMap<AccountMap>();
                    csv.WriteRecords(accounts);
                }

                Logger.Info("Finished exporting Chart Of Accounts");
            } catch (Exception ex) {
                Logger.Error("Failed to export Chart of Accounts!");
                Logger.Error($"Error: {ex.Message}");
                throw;
            }
        }

        private class Account {
            public string Company           { get; set; }
            public string GL_Account_Number { get; set; }
            public string GL_Account_Label  { get; set; }
            public string Classification    { get; set; }
            public string Action            { get; set; }
        }

        /// <summary>
        /// This determines the column names and order.
        /// </summary>
        private class AccountMap : ClassMap<Account> {
            public AccountMap() {
                Map(m => m.Company).Index(0).Name("Company");
                Map(m => m.GL_Account_Number).Index(1).Name("GL Account Number");
                Map(m => m.GL_Account_Label).Index(2).Name("GL Account Label");
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