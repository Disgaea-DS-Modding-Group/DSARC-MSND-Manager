namespace Properties
{
    internal static class Resources
    {
        public const string Filter_DatMsnd = "DAT/MSND Files (*.dat;*.msnd)|*.dat;*.msnd";
        public const string Status_MSND_Saved = "MSND saved";
        public const string Status_DSARC_Saved = "DSARC saved";
        public const string Status_FolderImported = "Folder imported";
        public const string Status_ExtractComplete = "Extract complete";
        public const string Status_FileReplaced = "File replaced.";
        public const string Status_FileReplacedUseSave = "File replaced. Use Save to rebuild DSARC/MSND.";
        public const string Status_ImportedAndStaged = "Imported and staged embedded archive (use Save to write).";
        public const string Status_NestedImportComplete = "Nested import complete";
        public const string Status_NestedExtractComplete = "Nested extract complete";
        public const string Status_ExtractedFormat = "{0} extracted";
        public const string Menu_ImportFolder = "Import Folder...";
        public const string Menu_ExtractAll = "Extract All...";
        public const string Menu_ExtractAllNested = "Extract All (Nested)...";
        public const string Menu_Extract = "Extract";
        public const string Menu_Replace = "Replace";
        public const string Dialog_SelectFolderToImport = "Select folder to import";
        public const string Dialog_ChooseOutputFolderForExtractAll = "Choose output folder for Extract All";
        public const string Dialog_SelectFolderToExtractItemTo = "Select folder to extract item to";
        public const string Dialog_SelectFolderToExtractChunkTo = "Select folder to extract chunk to";
        public const string Dialog_SelectRootFolderForExtraction = "Select root folder for extraction";
        public const string Dialog_SelectFolderContainingMsndParts = "Select folder containing .sseq/.sbnk/.swar to import";
        public const string Dialog_SelectFolderContainingExtractedArchive = "Select folder containing an extracted archive to import";
        public const string Dialog_SelectFolderContainingExtractedEmbeddedArchive = "Select folder containing the extracted embedded archive";
        public const string Dialog_SelectBaseFolderForNestedExtract = "Select base folder for nested extract";
        public const string Dialog_SelectBaseFolderForNestedExtractFromNode = "Select base folder for nested extract from node";
        public const string Warning_NoFilesFoundToImport = "No files found to import.";
        public const string Warning_OpenArchiveFirst = "Open an archive first.";
        public const string Warning_OpenDsarcFirst = "Open a DSARC archive first.";
        public const string Warning_InvalidSelection = "Invalid selection.";
        public const string Warning_InvalidParent = "Invalid parent";
        public const string Warning_InvalidChunk = "Invalid chunk";
        public const string Warning_SelectionNotEmbeddedArchive = "Selection is not an embedded archive.";
        public const string Warning_NestedImportDsarcOnly = "Nested import is for DSARC root only.";
        public const string Warning_ReplacementMustBeSseqSbnkSwar = "Replacement file must be .sseq, .sbnk, or .swar";
        public const string Error_FailedToRebuildRootArchive = "Failed to rebuild root archive.";
        public const string Error_FailedToRebuildArchive = "Failed to rebuild archive.";
        public const string Error_FileExistsCannotCreateFolder = "A file named '{0}' already exists in the selected folder; cannot create folder.";
    }
}
