﻿using System;
using System.Collections.Generic;
using System.IO;
using EpicorToYooz.Exports;
using EpicorToYooz.Properties;
using NLog;
using Renci.SshNet;

namespace EpicorToYooz {
    internal class Program {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private static void Main(string[] args) {
            try {
                Logger.Info("Clearing temp directory.");
                ClearTempDirectory();

                Logger.Info("Starting Epicor=>Yooz Exports.");

                ChartOfAccounts.Export();
                Vendors.Export();
                POs.Export();

                Logger.Info("Data exported.");
                Logger.Info("Beginning sFTP upload.");
                UploadFiles();

                Logger.Info("Finished!");
            } catch (Exception ex) {
                Logger.Fatal("Execution ended.");
                Logger.Fatal($"Error: {ex.Message}");
            }
            Console.ReadKey();
        }

        private static void UploadFiles() {
            // File name and destination sFTP path
            var files = new Dictionary<string, string> {
                {Settings.Default.FileName_ChartOfAccounts, Settings.Default.sFTP_Path_ChartOfAccount},
                {Settings.Default.FileName_Vendors, Settings.Default.sFTP_Path_Vendors},
                {Settings.Default.FileName_POs, Settings.Default.sFTP_Path_POs}
            };

            foreach (var file in files)
                UploadFile(file.Key,file.Value);
        }

        /// <summary>
        /// Upload a file to Yooz
        /// </summary>
        /// <param name="fileName">Name of file located in the temp directory.</param>
        /// <param name="remotePath">Folder upload should be placed in.</param>
        private static void UploadFile(string fileName, string remotePath) {
            var localPath     = Path.Combine(Settings.Default.TempDirectory, fileName);
            
            // Upload file as File-23423434.csv to avoid name conflicts
            var remoteFileName = $"{Path.GetFileNameWithoutExtension(fileName)}-{DateTime.Now.Ticks}.csv";

            Logger.Info($"Uploading {fileName} to ftp path: {remotePath}");
            var connectionInfo = new ConnectionInfo(Settings.Default.sFTP_URL,
                                                    Settings.Default.sFTP_Port,
                                                    Settings.Default.sFTP_User,
                                                    new PasswordAuthenticationMethod(Settings.Default.sFTP_User, Settings.Default.sFTP_Pass));

            using (var sFtp = new SftpClient(connectionInfo)) {
                sFtp.Connect();
                sFtp.ChangeDirectory(remotePath);
                using (var uploadStream = File.OpenRead(localPath))
                    sFtp.UploadFile(uploadStream, remoteFileName, false);

                sFtp.Disconnect();
            }

            Logger.Info("Upload complete.");
        }

        /// <summary>
        ///     Clear out old files in the temp directory. This is a safety measure to ensure we don't inadvertently upload old or
        ///     unwanted files.
        /// </summary>
        private static void ClearTempDirectory() {
            try {
                foreach (var filePath in Directory.GetFiles(Settings.Default.TempDirectory))
                    File.Delete(filePath);
            } catch (Exception ex) {
                Logger.Error("Failed to empty temp directory!");
                Logger.Error($"Error: {ex.Message}");
                throw;
            }
        }
    }
}