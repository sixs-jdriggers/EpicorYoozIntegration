using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Channels;
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

            // Create a new header if we haven't done so already.
            if (!invoiceHeadExists(service,exampleLine, groupID)) {
                // Get a new invoice dataset
                var invoiceDS = getInvoiceHead(groupID, service);

                // Populate our data and save the head.
                // After this we're ready to add lines.
                invoiceDS = populateInvoiceHead(invoiceDS, exampleLine, groupID, service);
            } 

            // Add lines to invoice
            ImportInvoiceLines(invoice, service);

            Logger.Info($"Invoice '{exampleLine.InvoiceNum}' created.");
        }
        
        private static void ImportInvoiceLines(YoozInvoice invoice, string service) {
            Logger.Info("Adding lines to invoice.");
            SetupREST();

            foreach (var invoiceLine in invoice.InvoiceLines) {
                Logger.Info($"Adding line for part: {invoiceLine.Description}");

                // Create the invoice lines with default values
                int.TryParse(invoiceLine.PONumber, out var poNum);
                dynamic postData;
                if (poNum!=0) {
                    Logger.Info("Creating receipt line.");
                    // Create receipt line
                    postData = getReceiptLines(service, invoiceLine);
                    selectAndInvoiceLines(service, postData, invoiceLine);

                    // Lookup our invoice with the newly created lines
                    postData = getInvoiceByID(service, invoiceLine);

                    // Set the quantity
                    postData = changeVendorQty(service, postData, invoiceLine);

                    // Set the price
                    if (invoiceLine.UnitPrice.HasValue)
                        postData = changeUnitCost(service, postData, invoiceLine);

                    // Save the final invoice
                    update(postData.parameters.ds, service);
                } else {
                    Logger.Info("Creating misc line.");
                    // Create Misc line
                    postData = getInvoiceByID(service, invoiceLine);
                    postData = createMiscLine(service, postData, invoiceLine);
                    postData = changePartNum_Misc(service, postData, invoiceLine);
                    postData = changeExtCost(service, invoiceLine, postData);

                    // Save the final invoice
                    update(postData.parameters.ds, service);
                }

                Logger.Info("Line added.");
            }

            Logger.Info("Finished adding lines.");
        }

        #region Header Creation
        private static dynamic getInvoiceHead(string groupID, string service) {
            Logger.Trace("Getting new invoice header.");
            var head = new {
                ds       = new { },
                cGroupID = groupID
            };

            return EpicorRest.DynamicPost(service, "GetNewAPInvHedInvoice", head);
        }

        private static dynamic populateInvoiceHead(dynamic invoiceDS, InvoiceLine exampleLine, string groupID, string service) {
            Logger.Trace("Poulating invoice header fields.");

            // Set PO Num or Vendor (For misc invoices)
            int.TryParse(exampleLine.PONumber, out var poNum);
            invoiceDS = poNum!=0 ? changeRefPONum(invoiceDS, exampleLine, service, poNum) : changeVendor(invoiceDS, exampleLine, service);

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

        private static dynamic changeVendor(dynamic invoiceDS, InvoiceLine exampleLine, string service) {
            // Set vendor
            var postData = new {
                invoiceDS.parameters.ds,
                ProposedVendorID = exampleLine.VendorID
            };

            return EpicorRest.DynamicPost(service, "ChangeVendorID", postData);
        }

        private static dynamic changeRefPONum(dynamic invoiceDS, InvoiceLine exampleLine, string service, int poNum) {
            // Set PO Number
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

        #region Receipt Line Creation
        // Process PO/Receipt Lines
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
            // Select lines for invoicing
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

            // Apply our selections
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

            // Add the selections to our invoice
            postData = new {
                postData.parameters.ds,
                opLOCMsg = ""
            };

            EpicorRest.DynamicPost(service, "InvoiceSelectedLines", postData);
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
                invoiceDS.ds,
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
                invoiceDS.ds,
                ProposedUnitCost = invoiceLine.UnitPrice
            };

            invoiceDS = EpicorRest.DynamicPost(service, "ChangeUnitCost", paramDS);
            return invoiceDS;
        }
        #endregion

        #region Misc Line Creation
        // Process Misc Lines
        private static object createMiscLine(string service, dynamic postData, InvoiceLine invoiceLine) {
            // Get new non-POLine
            postData = new {
                ds          = postData.ds,
                iVendorNum  = getVendorNumByID(invoiceLine.VendorID),
                cInvoiceNum = invoiceLine.InvoiceNum
            };

            postData = EpicorRest.DynamicPost(service, "GetNewAPInvDtlMiscellaneous", postData);

            return postData;
        }

        private static dynamic changePartNum_Misc(string service, dynamic postData, InvoiceLine invoiceLine) {
            // Populate line info
            var ds    = postData.parameters.ds;
            var lines = ds.APInvDtl.Count;
            ds.APInvDtl[lines-1].PartNum     = !string.IsNullOrWhiteSpace(invoiceLine.ProductCode) ? invoiceLine.ProductCode : invoiceLine.Description;
            ds.APInvDtl[lines-1].Description = invoiceLine.Description;

            postData = new {
                ProposedPartNum = !string.IsNullOrWhiteSpace(invoiceLine.ProductCode) ? invoiceLine.ProductCode : invoiceLine.Description,
                ds
            };

            return EpicorRest.DynamicPost(service, "ChangePartNum", postData);
        }

        private static dynamic changeExtCost(string service, InvoiceLine invoiceLine, dynamic ds) {
            ds = ds.parameters.ds;
            var lines = ds.APInvDtl.Count;

            ds.APInvDtl[lines-1].DocExtCost = invoiceLine.InvoiceAmount;

            var postData = new {
                ProposedExtCost = invoiceLine.InvoiceAmount,
                ds
            };

            return EpicorRest.DynamicPost(service, "ChangeExtCost", postData);
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

        private static dynamic getInvoiceByID(string service, InvoiceLine invoiceLine) {
            dynamic postData;
            // Look up the invoice with our newly created lines
            postData = new {
                vendorNum  = getVendorNumByID(invoiceLine.VendorID),
                invoiceNum = invoiceLine.InvoiceNum
            };

            postData = EpicorRest.DynamicPost(service, "GetByID", postData);
            return postData;
        }

        private static bool invoiceHeadExists(string service, InvoiceLine exampleLine, string groupId) {
            Logger.Info($"Looking up invoice '{exampleLine.InvoiceNum}'.");
            // Example call
            // https://epdev/epicor10/api/v1/Erp.BO.APInvoiceSvc/APInvoices?%24select=InvoiceNum%2C%20GroupID&%24filter=InvoiceNum%20eq%20'PBURNS-1-6'%20and%20GroupID%20eq%20'AM11216'
            var parameters = new Dictionary<string,string>() {
                {"$select","InvoiceNum,GroupID" },
                {"$filter", $"InvoiceNum eq '{exampleLine.InvoiceNum}' and GroupID eq '{groupId}'"}
            };

            var invoices = EpicorRest.DynamicGet(service, "APInvoices", parameters);
            var invoiceExists = invoices.value.Count > 0;
            Logger.Info($"Invoice exists: {invoiceExists}.");
            return invoiceExists;
        }
        #endregion
    }
}