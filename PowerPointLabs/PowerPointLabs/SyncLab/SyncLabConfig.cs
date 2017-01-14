using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using PowerPointLabs.FunctionalTestInterface.Impl;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs
{
    internal class SyncLabConfig
    {
#pragma warning disable 0618
        private const string DefaultSyncMasterFolderName = @"\PowerPointLabs Custom Sync";
        private const string DefaultSyncCategoryName = "My Sync";
        private const string SyncRootFolderConfigFileName = "SyncRootFolder.config";

        private readonly string _defaultSyncMasterFolderPrefix =
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        private string _configFilePath;

        # region Properties
        public string SyncRootFolder { get; set; }
        public string DefaultCategory { get; set; }
        # endregion

        # region Constructor
        public SyncLabConfig(string appDataFolder)
        {
            if (!PowerPointLabsFT.IsFunctionalTestOn)
            {
                SyncRootFolder = _defaultSyncMasterFolderPrefix + DefaultSyncMasterFolderName;
                DefaultCategory = DefaultSyncCategoryName;

                ReadShapeLabConfig(appDataFolder);
            }
            else
            {
                // if it's in FT, use new temp shape root folder every time
                var tmpPath = TempPath.GetTempTestFolder();
                var hash = DateTime.Now.GetHashCode();
                SyncRootFolder = tmpPath + DefaultSyncMasterFolderName + hash;
                DefaultCategory = DefaultSyncCategoryName + hash;
                _configFilePath = tmpPath + "ShapeRootFolder" + hash;
            }
        }
        # endregion

        # region Destructor
        ~SyncLabConfig()
        {
            // flush shape root folder & default category info to the file
            using (var fileWriter = File.CreateText(_configFilePath))
            {
                fileWriter.WriteLine(SyncRootFolder);
                fileWriter.WriteLine(DefaultCategory);
                
                fileWriter.Close();
            }
        }
        # endregion

        # region Helper Functions
        private void ReadShapeLabConfig(string appDataFolder)
        {
            _configFilePath = Path.Combine(appDataFolder, SyncRootFolderConfigFileName);

            if (File.Exists(_configFilePath) &&
                (new FileInfo(_configFilePath)).Length != 0)
            {
                using (var reader = new StreamReader(_configFilePath))
                {
                    SyncRootFolder = reader.ReadLine();
                    
                    // if we have a default category setting
                    if (reader.Peek() != -1)
                    {
                        DefaultCategory = reader.ReadLine();
                    }

                    reader.Close();
                }
            }
        }
        # endregion
    }
}