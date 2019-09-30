using System;
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
    internal class POs {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void Export() {
            try {
                Logger.Debug("Setting up Epicor REST library...");
                EpicorRest.AppPoolHost      = Settings.Default.Epicor_Server;
                EpicorRest.AppPoolInstance  = Settings.Default.Epicor_Instance;
                EpicorRest.UserName         = Settings.Default.Epicor_User;
                EpicorRest.Password         = Settings.Default.Epicor_Pass;
                EpicorRest.CallSettings = new CallSettings(Settings.Default.Epicor_Company, string.Empty, string.Empty, string.Empty);
                EpicorRest.IgnoreCertErrors = true;

                Logger.Debug($"Calling BAQ {Settings.Default.BAQ_POs}...");
                var jsonData = EpicorRest.GetBAQResultJSON(Settings.Default.BAQ_POs, null);
                var data     = JObject.Parse(jsonData)["value"].ToObject<DataTable>();

                Logger.Debug("Converting BAQ results to classes for export...");
                var orders = data.Rows.Cast<DataRow>()
                                 .Select(row => new PO {
                                     Action          = row[BAQColumns.Calculated_Action.ToString()].ToString(),
                                     VendorCode      = row[BAQColumns.Vendor_VendorID.ToString()].ToString(),
                                     VendorName      = row[BAQColumns.Vendor_Name.ToString()].ToString(),
                                     OrderNumber     = row[BAQColumns.POHeader_PONum.ToString()].ToString(),
                                     OrderDate       = row[BAQColumns.POHeader_OrderDate.ToString()].ToString(),
                                     Amount          = row[BAQColumns.POHeader_TotalOrder.ToString()].ToString(),
                                     AmountExlTax    = row[BAQColumns.Calculated_AmountExlTax.ToString()].ToString(),
                                     Currency        = row[BAQColumns.POHeader_CurrencyCode.ToString()].ToString(),
                                     OrderCreator    = row[BAQColumns.UserFile_Name.ToString()].ToString(),
                                     OrderApprover   = row[BAQColumns.PurAgent_BuyerID.ToString()].ToString(),
                                     Status          = row[BAQColumns.Calculated_Status.ToString()].ToString(),
                                     ItemNumber      = row[BAQColumns.PODetail_PartNum.ToString()].ToString(),
                                     ItemCode        = row[BAQColumns.Calculated_ItemCode.ToString()].ToString(),
                                     ItemDescription = row[BAQColumns.PODetail_LineDesc.ToString()].ToString(),
                                     ItemUnitPrice   = row[BAQColumns.PODetail_UnitCost.ToString()].ToString(),
                                     QuantityOrdered = row[BAQColumns.PODetail_OrderQty.ToString()].ToString()
                                 })
                                 .ToList();

                Logger.Info("Writing Purchase Order CSV file...");
                using (var writer = new StreamWriter($"{Settings.Default.TempDirectory}\\{Settings.Default.FileName_POs}"))
                using (var csv = new CsvWriter(writer)) {
                    csv.Configuration.HasHeaderRecord = false;
                    csv.Configuration.Delimiter       = "\t";
                    csv.Configuration.RegisterClassMap<POMap>();
                    csv.WriteComment(Settings.Default.Export_PO_Version);
                    csv.WriteComment(Settings.Default.Export_PO_Header);
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
            public string Action          { get; set; }
            public string VendorCode      { get; set; }
            public string VendorName      { get; set; }
            public string OrderNumber     { get; set; }
            public string OrderDate       { get; set; }
            public string Amount          { get; set; }
            public string AmountExlTax    { get; set; }
            public string Currency        { get; set; }
            public string OrderCreator    { get; set; }
            public string OrderApprover   { get; set; }
            public string Status          { get; set; }
            public string ItemNumber      { get; set; }
            public string ItemCode        { get; set; }
            public string ItemDescription { get; set; }
            public string ItemUnitPrice   { get; set; }
            public string QuantityOrdered { get; set; }
        }

        /// <summary>
        ///     This determines the column names and order.
        /// </summary>
        private class POMap : ClassMap<PO> {
            public POMap() {
                Map(m => m.Action).Index(0).Name("Action");
                Map(m => m.VendorCode).Index(1).Name("VendorCode");
                Map(m => m.VendorName).Index(2).Name("VendorName");
                Map(m => m.OrderNumber).Index(3).Name("OrderNumber");
                Map(m => m.OrderDate).Index(4).Name("OrderDate");
                Map(m => m.Amount).Index(5).Name("Amount");
                Map(m => m.AmountExlTax).Index(6).Name("AmountExlTax");
                Map(m => m.Currency).Index(7).Name("Currency");
                Map(m => m.OrderCreator).Index(8).Name("OrderCreator");
                Map(m => m.OrderApprover).Index(9).Name("OrderApprover");
                Map(m => m.Status).Index(10).Name("Status");
                Map(m => m.ItemNumber).Index(11).Name("ItemNumber");
                Map(m => m.ItemCode).Index(12).Name("ItemCode");
                Map(m => m.ItemDescription).Index(13).Name("ItemDescription");
                Map(m => m.ItemUnitPrice).Index(14).Name("ItemUnitPrice");
                Map(m => m.QuantityOrdered).Index(15).Name("QuantityOrdered");
            }
        }

        private enum BAQColumns {
            Calculated_Action,
            Vendor_VendorID,
            Vendor_Name,
            POHeader_PONum,
            POHeader_OrderDate,
            POHeader_TotalOrder,
            Calculated_AmountExlTax,
            POHeader_CurrencyCode,
            UserFile_Name,
            PurAgent_BuyerID,
            Calculated_Status,
            PODetail_PartNum,
            Calculated_ItemCode,
            PODetail_LineDesc,
            PODetail_UnitCost,
            PODetail_OrderQty
        }
    }
}