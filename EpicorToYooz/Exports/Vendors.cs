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
    internal class Vendors {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void Export() {
            try {
                Logger.Debug("Setting up Epicor REST library...");
                EpicorRest.AppPoolHost     = Settings.Default.Epicor_Server;
                EpicorRest.AppPoolInstance = Settings.Default.Epicor_Instance;
                EpicorRest.UserName        = Settings.Default.Epicor_User;
                EpicorRest.Password        = Settings.Default.Epicor_Pass;

                Logger.Debug($"Calling BAQ {Settings.Default.BAQ_Vendors}...");
                var data = EpicorRest.GetBAQResults(Settings.Default.BAQ_Vendors, null);

                Logger.Debug("Converting BAQ results to classes for export...");
                var vendors = data.Rows.Cast<DataRow>()
                                  .Select(row => new Vendor {
                                      Company           = row[BAQColumns.Vendor_Company.ToString()].ToString(),
                                      Third_Party_Code  = row[BAQColumns.Vendor_VendorID.ToString()].ToString(),
                                      Third_Party_Name  = row[BAQColumns.Vendor_Name.ToString()].ToString(),
                                      USA_EIN_or_TIN    = row[BAQColumns.Vendor_TaxPayerID.ToString()].ToString(),
                                      Phone_Number      = row[BAQColumns.Vendor_PhoneNum.ToString()].ToString(),
                                      Fax_Number        = row[BAQColumns.Vendor_FaxNum.ToString()].ToString(),
                                      Website           = row[BAQColumns.Vendor_VendURL.ToString()].ToString(),
                                      Address           = row[BAQColumns.Vendor_Address1.ToString()].ToString(),
                                      Address2          = row[BAQColumns.Vendor_Address2.ToString()].ToString(),
                                      Zip_Code          = row[BAQColumns.Vendor_ZIP.ToString()].ToString(),
                                      City              = row[BAQColumns.Vendor_City.ToString()].ToString(),
                                      Country_ISO_Code  = row[BAQColumns.Country_ISOCode.ToString()].ToString(),
                                      State_Code        = row[BAQColumns.Vendor_State.ToString()].ToString(),
                                      Contact_Name      = row[BAQColumns.VendCnt_Name.ToString()].ToString(),
                                      Contact_EMail     = row[BAQColumns.VendCnt_EmailAddress.ToString()].ToString(),
                                      Contact_Job_Title = row[BAQColumns.VendCnt_ContactTitle.ToString()].ToString()
                                  })
                                  .ToList();

                Logger.Info("Writing Vendors to CSV file...");
                using (var writer = new StreamWriter($"{Settings.Default.TempDirectory}\\{Settings.Default.FileName_Vendors}"))
                using (var csv = new CsvWriter(writer)) {
                    csv.Configuration.RegisterClassMap<VendorMap>();
                    csv.WriteRecords(vendors);
                }

                Logger.Info("Finished exporting Vendors.");
            } catch (Exception ex) {
                Logger.Error("Failed to export Vendors!");
                Logger.Error($"Error: {ex.Message}");
                throw;
            }
        }

        private class Vendor {
            public string Company           { get; set; }
            public string Third_Party_Code  { get; set; }
            public string Third_Party_Name  { get; set; }
            public string USA_EIN_or_TIN    { get; set; }
            public string Phone_Number      { get; set; }
            public string Fax_Number        { get; set; }
            public string Website           { get; set; }
            public string Address           { get; set; }
            public string Address2          { get; set; }
            public string Zip_Code          { get; set; }
            public string City              { get; set; }
            public string Country_ISO_Code  { get; set; }
            public string State_Code        { get; set; }
            public string Contact_Name      { get; set; }
            public string Contact_EMail     { get; set; }
            public string Contact_Job_Title { get; set; }
        }

        /// <summary>
        /// This determines the column names and order.
        /// </summary>
        private class VendorMap : ClassMap<Vendor> {
            public VendorMap() {
                Map(m => m.Company).Index(0).Name("Company");
                Map(m => m.Third_Party_Code).Index(1).Name("Third_Party_Code");
                Map(m => m.Third_Party_Name).Index(2).Name("Third_Party_Name");
                Map(m => m.USA_EIN_or_TIN).Index(3).Name("USA_EIN_or_TIN");
                Map(m => m.Phone_Number).Index(4).Name("Phone_Number");
                Map(m => m.Fax_Number).Index(5).Name("Fax_Number");
                Map(m => m.Website).Index(6).Name("Website");
                Map(m => m.Address).Index(7).Name("Address");
                Map(m => m.Address2).Index(8).Name("Address2");
                Map(m => m.Zip_Code).Index(9).Name("Zip_Code");
                Map(m => m.City).Index(10).Name("City");
                Map(m => m.Country_ISO_Code).Index(11).Name("Country_ISO_Code");
                Map(m => m.State_Code).Index(12).Name("State_Code");
                Map(m => m.Contact_Name).Index(13).Name("Contact_Name");
                Map(m => m.Contact_EMail).Index(14).Name("Contact_EMail");
                Map(m => m.Contact_Job_Title).Index(15).Name("Contact_Job_Title");
            }
        }

        private enum BAQColumns {
                Vendor_Company,
                Vendor_VendorID,
                Vendor_Name,
                Vendor_TaxPayerID,
                Vendor_PhoneNum,
                Vendor_FaxNum,
                Vendor_VendURL,
                Vendor_Address1,
                Vendor_Address2,
                Vendor_ZIP,
                Vendor_City,
                Country_ISOCode,
                Vendor_State,
                VendCnt_Name,
                VendCnt_EmailAddress,
                VendCnt_ContactTitle
            }
        }
    }