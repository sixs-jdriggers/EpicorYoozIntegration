using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;

namespace YoozToEpicor.Downloads {
    internal class YoozInvoice {
        /// <summary>
        ///     Read a CSV file and parse it into an invoice
        /// </summary>
        /// <param name="filePath">Path to CSV</param>
        public YoozInvoice(string filePath) {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader)) {
                csv.Configuration.RegisterClassMap<InvoiceLineMap>();
                InvoiceLines.AddRange(csv.GetRecords<InvoiceLine>());
            }
        }

        public YoozInvoice(IEnumerable<InvoiceLine> lines) {
            InvoiceLines = lines.ToList();
        }

        public readonly List<InvoiceLine> InvoiceLines = new List<InvoiceLine>();
    }

    public class InvoiceLine {
        public string    Entity        { get; set; }
        public string    VendorID      { get; set; }
        public DateTime? PostDate      { get; set; }
        public DateTime? InvoiceDate   { get; set; }
        public DateTime? DueDate       { get; set; }
        public string    GLAccount     { get; set; }
        public string    InvoiceNum    { get; set; }
        public decimal   InvoiceAmount { get; set; }
        public string    Description   { get; set; }
        public string    CostCenter    { get; set; }
        public string    SubAccount    { get; set; }
        public string    PONumber      { get; set; }
        public string    ProductCode   { get; set; }
        public string    ProductLabel  { get; set; }
        public decimal?  InvoiceQty    { get; set; }
        public decimal?  UnitPrice     { get; set; }
        public int?      POLine        { get; set; }
    }

    // Map CSV headers to class properties.
    // Be mindful, these are case-sensitive and Yooz is kinda sloppy about their casing.
    public sealed class InvoiceLineMap : ClassMap<InvoiceLine> {
        public InvoiceLineMap() {
            Map(m => m.Entity).Name("Entity");
            Map(m => m.VendorID).Name("Vendor ID");
            Map(m => m.PostDate).Name("Post Date");
            Map(m => m.InvoiceDate).Name("Invoice Date");
            Map(m => m.DueDate).Name("Due Date");
            Map(m => m.GLAccount).Name("G/L Account");
            Map(m => m.InvoiceNum).Name("Invoice Number");
            Map(m => m.InvoiceAmount).Name("Invoice Amount");
            Map(m => m.Description).Name("Description");
            Map(m => m.CostCenter).Name("Cost Center");
            Map(m => m.SubAccount).Name("Sub Account");
            Map(m => m.PONumber).Name("PO#");
            Map(m => m.ProductCode).Name("Product code");
            Map(m => m.ProductLabel).Name("Product label");
            Map(m => m.InvoiceQty).Name("Invoiced quantity");
            Map(m => m.UnitPrice).Name("Unit price");
            Map(m => m.POLine).Name("PO line#");
        }
    }
}