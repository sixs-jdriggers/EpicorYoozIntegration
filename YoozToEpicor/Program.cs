﻿using System;
using System.IO;
using System.Linq;
using NLog;
using Renci.SshNet;
using YoozToEpicor.Downloads;
using YoozToEpicor.Properties;

namespace YoozToEpicor {
    internal class Program {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private static void Main(string[] args) {
            Logger.Info($"Temp Directory: {Settings.Default.TempDirectory}");

            /*
             * We have a config option to skip downloading new files. This
             * makes it easier to test with arbitrary files if Yooz doesn't have good test data
             * available on the sFTP site.
             */
            if (!Settings.Default.DevOption_SkipDownloadingNewFiles) {
                Logger.Info("Downloading Yooz files.");
                downloadFiles();
            } else {
                Logger.Info("Skipping download of Yooz files. (Configured option.)");}

            var filesToProcess = Directory.GetFiles(Settings.Default.TempDirectory);
            Logger.Info($"Processing {filesToProcess.Length} files.");

            // Process invoices
            var invoiceGroup = $"Y_{DateTime.Today.ToString("MMddyy")}";
            Epicor.CreateInvoiceGroup(invoiceGroup);
            foreach (var file in filesToProcess)
                processFile(file, invoiceGroup);

            // Archive/Delete processed files
            if (Settings.Default.DeleteTempFilesAfterProcessing)
                archiveFiles();
        }

        private static void archiveFiles() {
            Logger.Info($"Deleting processed files from: {Settings.Default.TempDirectory}");
            var files = Directory.GetFiles(Settings.Default.TempDirectory);
            foreach (var file in files) {
                Logger.Trace($"Deleting: {file}");
                File.Delete(file);
            }
        }

        private static void processFile(string file, string groupID) {
            try {
                Logger.Info($"Parsing file: {file}");
                var invoiceData = new YoozInvoice(file);

                // Are there multiple invoices in this file?
                var invoicesInFile = invoiceData.InvoiceLines.Select(i => i.InvoiceNum).Distinct();
                Logger.Info($"File contains {invoicesInFile.Count()} invoices: {string.Join(", ", invoicesInFile)}");

                // Extract the invoice from the file
                foreach (var invoiceNum in invoicesInFile)
                    try {
                        var invoice = new YoozInvoice(invoiceData.InvoiceLines.Where(i => i.InvoiceNum==invoiceNum).ToList());
                        Epicor.ImportInvoice(invoice, groupID);
                    } catch (Exception ex) {
                        Logger.Error($"Failed to process invoice '{invoiceNum}'. Error: {ex.Message}");
                    }
            } catch (Exception ex) {
                Logger.Error($"Failed to process {file}. Error: {ex.Message}");
            }
        }

        /// <summary>
        ///     Download all files from FTP and clear FTP folder. (Clearing is a configurable option.)
        /// </summary>
        private static void downloadFiles() {
            Logger.Info($"Downloading files from ftp path: {Settings.Default.sFTP_Path}");
            var connectionInfo = new ConnectionInfo(Settings.Default.sFTP_URL,
                                                    Settings.Default.sFTP_Port,
                                                    Settings.Default.sFTP_User,
                                                    new PasswordAuthenticationMethod(Settings.Default.sFTP_User, Settings.Default.sFTP_Pass));

            using (var sFtp = new SftpClient(connectionInfo)) {
                sFtp.Connect();
                sFtp.ChangeDirectory(Settings.Default.sFTP_Path);
                var files = sFtp.ListDirectory(Settings.Default.sFTP_Path);

                // We only want CSVs, and no hidden files.
                foreach (var file in files.Where(file => !file.Name.StartsWith(".") && file.Name.EndsWith(".csv"))) {
                    Logger.Trace($"Downloading: {file.Name}");
                    using (Stream localFile = File.OpenWrite(Path.Combine(Settings.Default.TempDirectory, file.Name)))
                        sFtp.DownloadFile(file.Name, localFile);

                    if (Settings.Default.DeleteFTPFilesAfterDownload) {
                        Logger.Trace($"Deleting: {file.Name}");
                        sFtp.DeleteFile(file.Name);
                    }
                }

                sFtp.Disconnect();
            }

            Logger.Info("Download complete.");
        }
    }
}