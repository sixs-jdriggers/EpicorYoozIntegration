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

        public static void UnlockInvoiceGroup(string invoiceGroup) {
            // Setup the REST api and select our BO.
            SetupREST();
            const string service = "Erp.BO.APInvGrpSvc";

            var postData = new {InGroupID = invoiceGroup};
            EpicorRest.DynamicPost(service, "UnlockGroup", postData);
            
            Logger.Info($"Unlocked group: {invoiceGroup}");
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
            if (!invoiceHeadExists(service, exampleLine, groupID)) {
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
                Logger.Info($"Adding line for part: {invoiceLine.ProductCode}");

                // Create the invoice lines with default values
                int.TryParse(invoiceLine.PONumber, out var poNum);
                if (poNum!=0)
                    processPOLine(service, invoiceLine);
                else
                    processMiscLine(service, invoiceLine);

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

            return EpicorRest.DynamicPost(service, "GetNewAPInvHedInvoice", head).parameters.ds;
        }

        private static dynamic populateInvoiceHead(dynamic invoiceDS, InvoiceLine exampleLine, string groupID, string service) {
            Logger.Trace("Poulating invoice header fields.");

            // Set PO Num or Vendor (For misc invoices)
            int.TryParse(exampleLine.PONumber, out var poNum);
            invoiceDS = poNum!=0 ? changeRefPONum(invoiceDS, service, poNum) : changeVendor(invoiceDS, exampleLine, service);

            // Set Invoice Date
            invoiceDS = changeInvoiceDateEx(invoiceDS, exampleLine, service);

            // Set Invoice Amt
            invoiceDS = changeInvoiceVendorAmt(invoiceDS, exampleLine, service);

            // Set other fields
            invoiceDS.APInvHed[0].InvoiceNum = exampleLine.InvoiceNum;
            invoiceDS.APInvHed[0].DocumentID_c = exampleLine.DocumentID;

            // Make sure we don't override default dates with null values.
            if (exampleLine.InvoiceDate.HasValue)
                invoiceDS.APInvHed[0].InvoiceDate = exampleLine.InvoiceDate.Value;

            if (exampleLine.DueDate.HasValue)
                invoiceDS.APInvHed[0].DueDate = exampleLine.DueDate.Value;
            
            // Save the final invoice
            invoiceDS = update(invoiceDS, service);
            return invoiceDS;
        }

        private static dynamic changeVendor(dynamic invoiceDS, InvoiceLine exampleLine, string service) {
            // Set vendor
            var postData = new {
                ds               = invoiceDS,
                ProposedVendorID = exampleLine.VendorID
            };

            return EpicorRest.DynamicPost(service, "ChangeVendorID", postData).parameters.ds;
        }

        private static dynamic changeRefPONum(dynamic invoiceDS, string service, int poNum) {
            // Set PO Number
            invoiceDS.APInvHed[0].REFPONum = poNum;
            var changePOData = new {
                ProposedRefPONum = poNum,
                ds               = invoiceDS
            };

            return EpicorRest.DynamicPost(service, "ChangeRefPONum", changePOData).parameters.ds;
        }

        private static dynamic changeInvoiceDateEx(dynamic invoiceDS, InvoiceLine exampleLine, string service) {
            var postData = new {
                ds                  = invoiceDS,
                ProposedInvoiceDate = exampleLine.InvoiceDate ?? DateTime.Today,
                recalcAmts          = "",
                cMessageText        = ""
            };

            return EpicorRest.DynamicPost(service, "ChangeInvoiceDateEx", postData).parameters.ds;
        }

        private static dynamic changeInvoiceVendorAmt(dynamic invoiceDS, InvoiceLine exampleLine, string service) {
            var postData = new {
                ds                       = invoiceDS,
                ProposedInvoiceVendorAmt = exampleLine.InvoiceAmount
            };

            return EpicorRest.DynamicPost(service, "ChangeInvoiceVendorAmt", postData).parameters.ds;
        }

        private static dynamic update(dynamic invoiceDS, string service) {
            var updateData = new {ds = invoiceDS};

            return EpicorRest.DynamicPost(service, "Update", updateData).parameters.ds;
        }
        #endregion

        #region PO Line Creation
        // Process PO/Receipt Lines
        private static void processPOLine(string service, InvoiceLine invoiceLine) {
            Logger.Info("Creating PO invoice line.");
            // Retrieve receipt lines to see if PO was recieved. 
            var receiptLineDS = getReceiptLines(service, invoiceLine);

            // Was the line received? If not, create an unreceived invoice line
            if (lineWasReceived(invoiceLine, receiptLineDS))
                createReceiptLine(service, receiptLineDS, invoiceLine);
            else
                createUnreceivedLine(service, invoiceLine);

            // Lookup our invoice with the newly created lines
            var invoiceDS = getInvoiceByID(service, invoiceLine);

            // Set the quantity
            invoiceDS = changeVendorQty(service, invoiceDS, invoiceLine);

            // Set the price
            if (invoiceLine.UnitPrice.HasValue)
                invoiceDS = changeUnitCost(service, invoiceDS, invoiceLine);

            // Save the invoice
            invoiceDS = update(invoiceDS, service);
            // Set the GL Account
            invoiceDS = setGLAccount(service, invoiceLine, invoiceDS);
            // Save the changes 
            invoiceDS = update(invoiceDS, service);
        }

        private static dynamic getReceiptLines(string service, InvoiceLine invoiceLine) {
            Logger.Info("Checking for PO receipts.");
            // Get receipts
            dynamic postData = new {
                ds           = new { },
                InVendorNum  = getVendorNumByID(invoiceLine.VendorID),
                InInvoiceNum = invoiceLine.InvoiceNum,
                InPONum      = invoiceLine.PONumber
            };

            postData = EpicorRest.DynamicPost(service, "GetAPUninvoicedReceipts", postData);
            return postData.parameters.ds;
        }

        private static bool lineWasReceived(InvoiceLine invoiceLine, dynamic receiptLineDS) {
            foreach (var uninvoicedRcptLine in receiptLineDS.APUninvoicedRcptLines)
                if (uninvoicedRcptLine.POLine==invoiceLine.POLine)
                    return true; // We found a receipt!

            return false; // No receipt found!
        }

        private static void createReceiptLine(string service, dynamic receiptLineDS, InvoiceLine invoiceLine) {
            Logger.Info("Creating receipt line.");

            // Select lines for invoicing
            var purPoint = "";
            var packSlip = "";
            foreach (var uninvoicedRcptLine in receiptLineDS.APUninvoicedRcptLines) {
                if (uninvoicedRcptLine.POLine!=invoiceLine.POLine)
                    continue; // Wrong line

                uninvoicedRcptLine.SelectLine = true;
                uninvoicedRcptLine.RowMod     = "U";
                purPoint                      = uninvoicedRcptLine.PurPoint;
                packSlip                      = uninvoicedRcptLine.PackSlip;
                break; // Found our line!
            }

            // Apply our selections
            receiptLineDS = new {
                ds          = receiptLineDS,
                InVendorNum = getVendorNumByID(invoiceLine.VendorID),
                InPurPoint  = purPoint,
                InPONum     = invoiceLine.PONumber,
                InPackSlip  = packSlip,
                InDropShip  = false,
                invoiceLine.InvoiceNum,
                InGRNIClearing = false
            };

            receiptLineDS = EpicorRest.DynamicPost(service, "SelectUninvoicedRcptLines", receiptLineDS).parameters.ds;

            // Add the selections to our invoice
            receiptLineDS = new {
                ds       = receiptLineDS,
                opLOCMsg = ""
            };

            EpicorRest.DynamicPost(service, "InvoiceSelectedLines", receiptLineDS);
        }

        private static void createUnreceivedLine(string service, InvoiceLine invoiceLine) {
            Logger.Info("Creating unreceived line");
            // Get our invoice.
            var invoiceDS = getInvoiceByID(service, invoiceLine);

            // Create a new line
            dynamic postData = new {
                ds          = invoiceDS,
                iVendorNum  = getVendorNumByID(invoiceLine.VendorID),
                cInvoiceNum = invoiceLine.InvoiceNum
            };

            invoiceDS = EpicorRest.DynamicPost(service, "GetNewAPInvDtlUnreceived", postData).parameters.ds;

            // Set the PO
            postData = new {
                ds            = invoiceDS,
                ProposedPONum = invoiceLine.PONumber
            };

            invoiceDS = EpicorRest.DynamicPost(service, "ChangePONum", postData).parameters.ds;

            // Set the PO Line
            postData = new {
                ds             = invoiceDS,
                ProposedPOLine = invoiceLine.POLine
            };

            invoiceDS = EpicorRest.DynamicPost(service, "ChangePOLine", postData).parameters.ds;

            // Save the line
            update(invoiceDS, service);
        }

        private static dynamic changeVendorQty(string service, dynamic invoiceDS, InvoiceLine invoiceLine) {
            foreach (var invDtl in invoiceDS.APInvDtl) {
                if (invDtl.POLine!=invoiceLine.POLine)
                    continue; // Wrong line

                invDtl.RowMod    = "U";
                invDtl.VendorQty = invoiceLine.InvoiceQty;

                break; // Done with our line!
            }

            var paramDS = new {
                ds                = invoiceDS,
                ProposedVendorQty = invoiceLine.InvoiceQty
            };

            invoiceDS = EpicorRest.DynamicPost(service, "ChangeVendorQty", paramDS);
            return invoiceDS.parameters.ds;
        }

        private static dynamic changeUnitCost(string service, dynamic invoiceDS, InvoiceLine invoiceLine) {
            foreach (var invDtl in invoiceDS.APInvDtl) {
                if (invDtl.POLine!=invoiceLine.POLine)
                    continue; // Wrong line

                invDtl.UnitCost = invoiceLine.UnitPrice;
                invDtl.RowMod   = "U";
                break; // Done with our line
            }

            var paramDS = new {
                ds               = invoiceDS,
                ProposedUnitCost = invoiceLine.UnitPrice
            };

            invoiceDS = EpicorRest.DynamicPost(service, "ChangeUnitCost", paramDS);
            return invoiceDS.parameters.ds;
        }
        #endregion

        #region Misc Line Creation
        // Process Misc Lines
        private static void processMiscLine(string service, InvoiceLine invoiceLine) {
            Logger.Info("Creating misc invoice line.");
            // Create Misc line
            var invoiceDS = getInvoiceByID(service, invoiceLine);
            invoiceDS = createMiscLine(service, invoiceDS, invoiceLine);
            invoiceDS = changePartNum_Misc(service, invoiceDS, invoiceLine);
            invoiceDS = changeExtCost(service, invoiceLine, invoiceDS);
            // Save the invoice
            invoiceDS = update(invoiceDS, service);
            // Set the GL Account
            invoiceDS = setGLAccount(service, invoiceLine, invoiceDS);
            // Save the changes 
            invoiceDS = update(invoiceDS, service);
        }

        private static dynamic createMiscLine(string service, dynamic invoiceDS, InvoiceLine invoiceLine) {
            // Get new non-POLine
            invoiceDS = new {
                ds          = invoiceDS,
                iVendorNum  = getVendorNumByID(invoiceLine.VendorID),
                cInvoiceNum = invoiceLine.InvoiceNum
            };

            invoiceDS = EpicorRest.DynamicPost(service, "GetNewAPInvDtlMiscellaneous", invoiceDS);

            return invoiceDS.parameters.ds;
        }

        private static dynamic changePartNum_Misc(string service, dynamic invoiceDS, InvoiceLine invoiceLine) {
            // Populate line info
            var lines = invoiceDS.APInvDtl.Count;

            // Use part description if available
            // Else part num
            // Else invoice description.
            // All else fails, just use the invoice num.
            var description = !string.IsNullOrWhiteSpace(invoiceLine.ProductLabel) ? invoiceLine.ProductLabel :
                              !string.IsNullOrWhiteSpace(invoiceLine.ProductCode)  ? invoiceLine.ProductCode :
                              !string.IsNullOrWhiteSpace(invoiceLine.Description)  ? invoiceLine.Description : invoiceLine.InvoiceNum;

            // If we don't have the part, use the description
            var partNum = !string.IsNullOrWhiteSpace(invoiceLine.ProductCode) ? invoiceLine.ProductCode : description;
            invoiceDS.APInvDtl[lines-1].PartNum     = partNum;
            invoiceDS.APInvDtl[lines-1].Description = description;

            invoiceDS = new {
                ProposedPartNum = partNum,
                ds              = invoiceDS
            };

            return EpicorRest.DynamicPost(service, "ChangePartNum", invoiceDS).parameters.ds;
        }

        private static dynamic changeExtCost(string service, InvoiceLine invoiceLine, dynamic invoiceDS) {
            var lines = invoiceDS.APInvDtl.Count;

            invoiceDS.APInvDtl[lines-1].DocExtCost = invoiceLine.InvoiceAmount;

            var postData = new {
                ProposedExtCost = invoiceLine.InvoiceAmount,
                ds              = invoiceDS
            };

            return EpicorRest.DynamicPost(service, "ChangeExtCost", postData).parameters.ds;
        }

        private static dynamic setGLAccount(string service, InvoiceLine invoiceLine, dynamic invoiceDS) {
            // Find the APInvDtlExpTGLC record. 
            // This is the GL info for the invoice line
            var TGLCCount = invoiceDS.APInvExpTGLC.Count;

            invoiceDS.APInvExpTGLC[TGLCCount-1].GLAccount = $"{invoiceLine.GLAccount}|{invoiceLine.Entity}|{invoiceLine.CostCenter}";
            invoiceDS.APInvExpTGLC[TGLCCount-1].SegValue1 = invoiceLine.GLAccount;
            invoiceDS.APInvExpTGLC[TGLCCount-1].SegValue2 = invoiceLine.Entity;
            invoiceDS.APInvExpTGLC[TGLCCount-1].SegValue3 = invoiceLine.CostCenter;
            invoiceDS.APInvExpTGLC[TGLCCount-1].RowMod    = "U";

            // Set the account
            return invoiceDS;
        }
        #endregion

        #region Lookups and utility functions
        private static void SetupREST() {
            EpicorRest.AppPoolHost      = Settings.Default.Epicor_Server;
            EpicorRest.AppPoolInstance  = Settings.Default.Epicor_Instance;
            EpicorRest.UserName         = Settings.Default.Epicor_User;
            EpicorRest.Password         = Settings.Default.Epicor_Pass;
            EpicorRest.CallSettings     = new CallSettings(Settings.Default.Epicor_Company, string.Empty, string.Empty, string.Empty);
            EpicorRest.IgnoreCertErrors = true;
        }

        private static int getVendorNumByID(string vendorID) {
            Logger.Debug($"Looking up VendorNum for ID '{vendorID}'.");
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
            // Look up the invoice with our newly created lines
            dynamic invoiceDS = new {
                vendorNum  = getVendorNumByID(invoiceLine.VendorID),
                invoiceNum = invoiceLine.InvoiceNum
            };

            invoiceDS = EpicorRest.DynamicPost(service, "GetByID", invoiceDS);
            return invoiceDS.ds;
        }

        private static bool invoiceHeadExists(string service, InvoiceLine exampleLine, string groupId) {
            Logger.Info($"Looking up invoice '{exampleLine.InvoiceNum}'.");
            // Example call
            // https://epdev/epicor10/api/v1/Erp.BO.APInvoiceSvc/APInvoices?%24select=InvoiceNum%2C%20GroupID&%24filter=InvoiceNum%20eq%20'PBURNS-1-6'%20and%20GroupID%20eq%20'AM11216'
            var parameters = new Dictionary<string, string> {
                {"$select", "InvoiceNum,GroupID"},
                {"$filter", $"InvoiceNum eq '{exampleLine.InvoiceNum}' and GroupID eq '{groupId}'"}
            };

            var invoices      = EpicorRest.DynamicGet(service, "APInvoices", parameters);
            var invoiceExists = invoices.value.Count > 0;
            Logger.Info($"Invoice exists: {invoiceExists}.");
            return invoiceExists;
        }
        #endregion
    }
}