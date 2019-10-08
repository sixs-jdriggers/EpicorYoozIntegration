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
        ///     Upsert an AP invoice group
        /// </summary>
        /// <param name="groupID">Group to Upsert</param>
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

        /// <summary>
        ///     Create/update an AP invoice from Yooz data
        /// </summary>
        /// <param name="invoice"></param>
        public static void ImportInvoice(YoozInvoice invoice, string groupID) {
            Logger.Info($"Adding invoice '{invoice.InvoiceLines.First().InvoiceNum}' to group '{groupID}'.");

            // Setup the REST api and select our BO.
            SetupREST();
            var service = "Erp.BO.APInvoiceSvc";

            // Yooz uses a flat data format for invoices.
            // So we just grab the first line to get the header values.
            var exampleLine = invoice.InvoiceLines.First();

            // Get a new invoice dataset
            var invoiceDS = getInvoiceHead(groupID, service);

            // Populate our data and save the head.
            // After this we're ready to add lines.
            invoiceDS = populateInvoiceHead(invoiceDS, exampleLine, groupID, service);

            // Add lines to invoice
            ImportInvoiceLines(invoiceDS, invoice, groupID, service);

            Logger.Info($"Invoice '{exampleLine.InvoiceNum}' created.");
        }

        #region Header Creation
        private static dynamic getInvoiceHead(string groupID, string service) {
            Logger.Trace("Getting new invoice dataset.");
            var head = new {
                ds       = new { },
                cGroupID = groupID
            };

            return EpicorRest.DynamicPost(service, "GetNewAPInvHedInvoice", head);
        }

        private static dynamic populateInvoiceHead(dynamic invoiceDS, InvoiceLine exampleLine, string groupID, string service) {
            Logger.Trace("Poulating invoice header fields.");

            // Set PO Num
            invoiceDS = changeRefPONum(invoiceDS, exampleLine, service, out int poNum);

            // Set Invoice Date
            invoiceDS = changeInvoiceDateEx(invoiceDS, exampleLine, service);

            // Set Invoice Amt
            invoiceDS = changeInvoiceVendorAmt(invoiceDS, exampleLine, service);

            // Set other fields
            invoiceDS.parameters.ds.APInvHed[0].InvoiceNum = exampleLine.InvoiceNum;

            // Make sure we don't override default dates with null values.
            if (exampleLine.InvoiceDate.HasValue)
                invoiceDS.parameters.ds.APInvHed[0].InvoiceDate = exampleLine.InvoiceDate.Value;

            if (exampleLine.DueDate.HasValue)
                invoiceDS.parameters.ds.APInvHed[0].DueDate = exampleLine.DueDate.Value;

            // Save the final invoice
            invoiceDS = update(invoiceDS.parameters.ds, service);
            return invoiceDS;
        }

        private static dynamic changeRefPONum(dynamic invoiceDS, InvoiceLine exampleLine, string service, out int poNum) {
            // Set PO Number
            int.TryParse(exampleLine.PONumber, out poNum);
            invoiceDS.parameters.ds.APInvHed[0].REFPONum = poNum;
            var changePOData = new {
                ProposedRefPONum = poNum,
                invoiceDS.parameters.ds
            };

            return EpicorRest.DynamicPost(service, "ChangeRefPONum", changePOData);
        }

        private static dynamic changeInvoiceDateEx(dynamic invoiceDS, InvoiceLine exampleLine, string service) {
            var postData = new {
                invoiceDS.parameters.ds,
                ProposedInvoiceDate = exampleLine.InvoiceDate ?? DateTime.Today,
                recalcAmts          = "",
                cMessageText        = ""
            };

            return EpicorRest.DynamicPost(service, "ChangeInvoiceDateEx", postData);
        }

        private static dynamic changeInvoiceVendorAmt(dynamic invoiceDS, InvoiceLine exampleLine, string service) {
            var postData = new {
                invoiceDS.parameters.ds,
                ProposedInvoiceVendorAmt = exampleLine.InvoiceAmount
            };

            return EpicorRest.DynamicPost(service, "ChangeInvoiceVendorAmt", postData);
        }

        private static dynamic update(dynamic invoiceDS, string service) {
            var updateData = new {ds = invoiceDS};

            return EpicorRest.DynamicPost(service, "Update", updateData);
        }
        #endregion

        #region Line Creation
        private static void ImportInvoiceLines(dynamic invoiceDS, YoozInvoice invoice, string groupID, string service) {
            Logger.Info("Adding lines to invoice.");
            SetupREST();

            foreach (var invoiceLine in invoice.InvoiceLines) {
                Logger.Info($"Adding line for part: {invoiceLine.Description}");

                // Create the invoice lines with default values
                var postData = getReceiptLines(service, invoiceLine);
                selectAndInvoiceLines(service, postData, invoiceLine);

                // Lookup our invoice with the newly created lines
                postData = getByID(service, invoiceLine);

                // Set the quantity
                postData = changeVendorQty(service, postData, invoiceLine);

                // Set the price
                if (invoiceLine.UnitPrice.HasValue)
                    postData = changeUnitCost(service, postData, invoiceLine);
                
                // Save the final invoice
                postData = update(postData.parameters.ds, service);

                Logger.Info("Line added.");
            }

            Logger.Info("Finished adding lines.");
        }

        private static dynamic getReceiptLines(string service, InvoiceLine invoiceLine) {
            // Get receipts
            dynamic postData = new {
                ds           = new { },
                InVendorNum  = getVendorNumByID(invoiceLine.VendorID),
                InInvoiceNum = invoiceLine.InvoiceNum,
                InPONum      = invoiceLine.PONumber
            };

            postData = EpicorRest.DynamicPost(service, "GetAPUninvoicedReceipts", postData);
            return postData;
        }

        private static void selectAndInvoiceLines(string service, dynamic postData, InvoiceLine invoiceLine) {
            // Select lines
            var purPoint = "";
            var packSlip = "";
            foreach (var uninvoicedRcptLine in postData.parameters.ds.APUninvoicedRcptLines) {
                if (!string.Equals(uninvoicedRcptLine.PartNum.ToString(), invoiceLine.ProductCode, StringComparison.CurrentCultureIgnoreCase))
                    continue;

                uninvoicedRcptLine.SelectLine = true;
                uninvoicedRcptLine.RowMod     = "U";
                purPoint                      = uninvoicedRcptLine.PurPoint;
                packSlip                      = uninvoicedRcptLine.PackSlip;
                break; // Found our line!
            }

            postData = new {
                postData.parameters.ds,
                InVendorNum = getVendorNumByID(invoiceLine.VendorID),
                InPurPoint  = purPoint,
                InPONum     = invoiceLine.PONumber,
                InPackSlip  = packSlip,
                InDropShip  = false,
                invoiceLine.InvoiceNum,
                InGRNIClearing = false
            };

            postData = EpicorRest.DynamicPost(service, "SelectUninvoicedRcptLines", postData);

            // Invoice the lines
            postData = new {
                postData.parameters.ds,
                opLOCMsg = ""
            };

            // Invoice receipts
            EpicorRest.DynamicPost(service, "InvoiceSelectedLines", postData);
        }

        private static dynamic getByID(string service, InvoiceLine invoiceLine) {
            dynamic postData;
            // Look up the invoice with our newly created lines
            postData = new {
                vendorNum  = getVendorNumByID(invoiceLine.VendorID),
                invoiceNum = invoiceLine.InvoiceNum
            };

            postData = EpicorRest.DynamicPost(service, "GetByID", postData);
            return postData;
        }

        private static dynamic changeVendorQty(string service, dynamic invoiceDS, InvoiceLine invoiceLine) {
            foreach (var invDtl in invoiceDS.ds.APInvDtl) {
                if (!string.Equals(invDtl.PartNum.ToString(), invoiceLine.ProductCode, StringComparison.CurrentCultureIgnoreCase))
                    continue; // Wrong line

                invDtl.RowMod    = "U";
                invDtl.VendorQty = invoiceLine.InvoiceQty;

                break; // Done with our line!
            }

            var paramDS = new {
                ds                = invoiceDS.ds,
                ProposedVendorQty = invoiceLine.InvoiceQty
            };

            invoiceDS = EpicorRest.DynamicPost(service, "ChangeVendorQty", paramDS);
            return invoiceDS;
        }

        private static dynamic changeUnitCost(string service, dynamic invoiceDS, InvoiceLine invoiceLine) {
            foreach (var invDtl in invoiceDS.ds.APInvDtl) {
                if (!string.Equals(invDtl.PartNum.ToString(), invoiceLine.ProductCode, StringComparison.CurrentCultureIgnoreCase))
                    continue; // Wrong line

                invDtl.UnitCost = invoiceLine.UnitPrice;
                invDtl.RowMod   = "U";
                break; // Done with our line
            }

            var paramDS = new {
                ds               = invoiceDS.ds,
                ProposedUnitCost = invoiceLine.UnitPrice
            };

            invoiceDS = EpicorRest.DynamicPost(service, "ChangeUnitCost", paramDS);
            return invoiceDS;
        }
        #endregion

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

            var vendor = EpicorRest.DynamicGet("Erp.BO.VendorSvc", "Vendors", parameters);

            if (vendor.value.Count < 1)
                throw new Exception($"Unable to locate vendor by ID '{vendorID}'.");

            return vendor.value[0].VendorNum;
        }
        #endregion
    }
}