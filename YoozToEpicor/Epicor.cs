using System;
using System.Collections.Generic;
using System.Linq;
using EpicorRestAPI;
using NLog;
using YoozToEpicor.Downloads;
using YoozToEpicor.Properties;

namespace YoozToEpicor {
    internal class Epicor {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        ///     Create/update an AP invoice from Yooz data
        /// </summary>
        /// <param name="invoice"></param>
        public static void ImportInvoice(YoozInvoice invoice, string groupID) {
            Logger.Info($"Adding invoice '{invoice.InvoiceLines.First().InvoiceNum}' to group '{groupID}'.");

            // Setup the REST api and select our BO.
            SetupREST();
            var service    = "Erp.BO.APInvoiceSvc";

            // Upsert Header
            // Yooz uses a flat data format for invoices.
            // So we just grab the first line to get the header values.
            var exampleLine = invoice.InvoiceLines.First();
            var head = new {
                ds = new {
                    APInvHed = new {
                        Company = Settings.Default.Epicor_Company,
                        InvoiceNum = exampleLine.InvoiceNum,
                        InvoiceDate = exampleLine.InvoiceDate,
                        DueDate = exampleLine.DueDate,
                        vendorNum  = getVendorNumByID(exampleLine.VendorID),
                        InvoiceAmt = exampleLine.InvoiceAmount
                    }
                }
            };

            EpicorRest.DynamicPost(service, "GetNewAPInvHed", head);
            
            Logger.Info($"Invoice '{exampleLine.InvoiceNum}' created.");

            ImportInvoiceLines(invoice);
        }

        public static void CreateInvoiceGroup(string groupID) {
            // Setup the REST api and select our BO.
            SetupREST();
            const string service = "Erp.BO.APInvGrpSvc";
            const string method  = "APInvGrps";

            Logger.Info($"Looking for invoice group: {groupID}");

            // Does the group already exist?
            var parameters    = new Dictionary<string, string> {{"$filter", $"GroupID eq '{groupID}'"}};
            var existingGroup = EpicorRest.DynamicGet(service, method, parameters);
            if (existingGroup.value.Count > 0) {
                Logger.Info("Found group.");
                return; // Group already exists!
            }

            // Create if it doesn't exist
            Logger.Info("Creating group.");

            var group = new {
                Company = Settings.Default.Epicor_Company,
                GroupID = groupID
            };

            EpicorRest.DynamicPost(service, method, group);
        }

        private static void ImportInvoiceLines(YoozInvoice invoice) {
            Logger.Info("Adding lines to invoice.");
            SetupREST();
            var service      = "Erp.BO.APInvoiceSvc";
            var detailMethod = "APInvDtls";

            foreach (var invoiceLine in invoice.InvoiceLines) {
                var line = new { };
                EpicorRest.DynamicPost(service, detailMethod, line);
                Logger.Info($"Adding line for part: {invoiceLine.Description}");
            }
        }

        #region Lookups and utility functions
        private static void SetupREST() {
            Logger.Debug("Setting up Epicor REST library...");
            EpicorRest.AppPoolHost      = Settings.Default.Epicor_Server;
            EpicorRest.AppPoolInstance  = Settings.Default.Epicor_Instance;
            EpicorRest.UserName         = Settings.Default.Epicor_User;
            EpicorRest.Password         = Settings.Default.Epicor_Pass;
            EpicorRest.CallSettings     = new CallSettings(Settings.Default.Epicor_Company, string.Empty, string.Empty, string.Empty);
            EpicorRest.IgnoreCertErrors = true;
        }

        private static int getVendorNumByID(string vendorID) {
            Logger.Info($"Looking up VendorNum for ID '{vendorID}'.");
            SetupREST();
            var parameters = new Dictionary<string, string> {
                {"$filter", $"VendorID eq '{vendorID}'"},
                {"$select", "VendorNum,VendorID"}
            };
            var vendor     = EpicorRest.DynamicGet("Erp.BO.VendorSvc", "Vendors", parameters);

            if (vendor.value.Count < 1)
                throw new Exception($"Unable to locate vendor by ID '{vendorID}'.");

            return vendor.value[0].VendorNum;
        }
        #endregion
    }
}