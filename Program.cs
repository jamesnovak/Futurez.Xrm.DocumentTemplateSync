using System;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;

using Futurez.Entities;

namespace Futurez.Xrm.DocumentTemplateSync
{
    class DocumentSync
    {
        public const string DT_FORMAT = "yyyy-MM-dd HH.mm.ss";

        static readonly StringBuilder _logger = new StringBuilder();

        public static StringBuilder Logger { get => _logger; }

        public static void Main(string[] args)
        {
            // load the 
            var syncSettings = SyncSettings.FromFile(args);

            try {
                ReportProgress("DoExport!");
                DoExport(syncSettings);

                ReportProgress("DoImport!");
                DoImport(syncSettings);
            }
            catch (FaultException<OrganizationServiceFault> ex) {
                ReportProgress($"General exception has occurred: \n{ex.Message}");
            }

            // output the log file
            SaveLogFile().Wait();
        }

        /// <summary>
        /// Do the work to export the files 
        /// </summary>
        /// <param name="settings"></param>
        private static void DoExport(SyncSettings settings)
        {
            // create the export and import folders 
            if (!Directory.Exists(settings.ExportFolder)) {
                Directory.CreateDirectory(settings.ExportFolder);
            }

            // retrieve the templates from the source system if connection string has been set
            if (settings.ConnStringSource != null && !settings.ImportTargetOnly) 
            {
                ReportProgress($"Connecting to source environment: {settings.ConnStringSourceUrl}");

                using (var service = new CrmServiceClient(settings.ConnStringSource))
                {
                    if (service.OrganizationServiceProxy == null) {
                        throw new ApplicationException($"Unable to connect to CRM:{settings.ConnStringSourceUrl}");
                    }

                    // only grab content if we are not logging
                    var sourceTemplates = DocumentTemplateEdit.GetDocumentTemplates(service, settings, !settings.LoggingOnly);

                    foreach (var doc in sourceTemplates) 
                    {
                        var downloadFileName = Path.Combine(settings.ExportFolder, doc.FileName);

                        try {

                            ReportProgress($"File exported - FileName: {downloadFileName}, Entity: {doc.AssociatedEntityLogicalName}");

                            if (!settings.LoggingOnly) {

                                // make sure we export the latest!
                                if (File.Exists(downloadFileName)) {
                                    File.Delete(downloadFileName);
                                }
                                // counting on relatively small files with the single write 
                                using (var saveMe = new FileStream(downloadFileName, FileMode.OpenOrCreate, FileAccess.Write)) {
                                    saveMe.Write(doc.Content, 0, doc.Content.Length);
                                    saveMe.Close();
                                }
                            }
                        }
                        catch (Exception ex) {
                            ReportProgress($"An error occurred exporting {downloadFileName}\n{ex.Message}");
                        }
                    }
                }
            }
            else {
                var msg = "Skipping download:";
                if (settings.ConnStringSource == null) {
                    msg += $", No source environment connection string provided";
                }
                if (settings.ImportTargetOnly) {
                    msg += $", ImportTargetOnly specified";
                }
                ReportProgress(msg);
            }
        }

        /// <summary>
        /// Perform the work to import the files 
        /// </summary>
        /// <param name="settings"></param>
        private static void DoImport(SyncSettings settings)
        {
            // now that we have exported the files, import them
            if (settings.ConnStringTarget != null && !settings.ExportSourceOnly) 
            {
                // get all the files from the folder.  If we 
                var folder = (settings.ImportFolder != null) ? settings.ImportFolder : settings.ExportFolder;

                if (!Directory.Exists(folder)) {
                    ReportProgress($"Skipping import. Folder not found: {folder}");
                    return;
                }

                ReportProgress($"Connecting to target environment: {settings.ConnStringTargetUrl}");

                using (var service = new CrmServiceClient(settings.ConnStringTarget)) {

                    if (service.OrganizationServiceProxy == null) {
                        throw new ApplicationException($"Unable to connect to CRM:{settings.ConnStringTargetUrl}");
                    }

                    // first, retrieve the list of all doc templates in the source system 
                    // kind of defeats the purpose of an upsert, but need to check this for Create only.
                    var targetTemplates = DocumentTemplateEdit.GetDocumentTemplates(service);

                    // load the files from the folder 
                    var files = Directory.EnumerateFiles(folder);
                    foreach (var file in files) 
                    {
                        var fileName = Path.GetFileName(file);

                        // is this a new template or an update to an existing template?
                        var existingTemplate = targetTemplates.Where(d => d.FileName == fileName).FirstOrDefault();

                        // if the Create Only flag is set and the template exists, skip it.
                        if (settings.CreateOnly && (existingTemplate != null)) {
                            ReportProgress($"Skipping '{file}' because Update Only flag is set and the template exists on the target system: {existingTemplate.FileName}, Created On: {existingTemplate.CreatedOn}");
                        }
                        else 
                        {
                            // load up the file and do some checks... 
                            var fileUpload = new FileUpload(file, service);
                            var upsertId = Guid.NewGuid();

                            var description = $"New {fileUpload.TemplateType} template. Source file: {fileUpload.FileName}";

                            // check to see if the file has issues 
                            if (fileUpload.IsIgnored) {
                                ReportProgress($"Skipping file '{fileUpload.FileName}', Note: {fileUpload.Note}");
                            }
                            else {
                                // might be doing something wrong, but upsert only seems to work with the ID. need to investigate
                                if (existingTemplate == null) {
                                    // create us a new file
                                    description = $"Updated {fileUpload.TemplateType} template. Source file: {fileUpload.FileName}. Date: {DateTime.Now.ToShortDateString()} {DateTime.Now.ToShortTimeString()}";
                                }
                                else {
                                    upsertId = existingTemplate.Id;
                                }

                                if (!settings.LoggingOnly) {
                                    try {
                                        var target = new Entity("documenttemplate", upsertId);
                                        target.Attributes["name"] = fileUpload.TemplateName;
                                        target.Attributes["documenttype"] = new OptionSetValue(fileUpload.TemplateTypeValue);
                                        target.Attributes["description"] = description;
                                        target.Attributes["content"] = Convert.ToBase64String(fileUpload.FileContents);

                                        var upsert = new UpsertRequest() {
                                            Target = target
                                        };

                                        var response = service.Execute(upsert) as UpsertResponse;
                                        if (response.RecordCreated) {
                                            ReportProgress($"Created new template with file: '{fileUpload.FileName}'");
                                        }
                                        else {
                                            ReportProgress($"Updated template with file: '{fileUpload.FileName}'");
                                        }
                                    }
                                    catch (FaultException ex) {
                                        // seeing some issues with system Excel templates and the error ' The hidden sheet data is corrupted.'
                                        ReportProgress($"An error occured uploading the template: {fileUpload.FileName} -- {ex.Message}");
                                    }
                                }
                                else {
                                    ReportProgress($"Logging: {description}");
                                }
                            }
                        }
                    }
                }
            }
            else {
                var msg = "Skipping upload:";
                if (settings.ConnStringTarget == null) {
                    msg += $", No target environment connection string provided";
                }
                if (settings.ExportSourceOnly) {
                    msg += $", ExportSourceOnly specified";
                }
                ReportProgress(msg);
            }
        }

        /// <summary>
        /// Save the log!
        /// </summary>
        /// <returns></returns>
        static async Task SaveLogFile()
        {
            var now = DateTime.Now.ToString(DT_FORMAT);
            var logFile = $"PortalRecordsMoverApp {now}.log";

            byte[] encodedText = Encoding.Unicode.GetBytes(Logger.ToString());

            using (FileStream sourceStream = new FileStream(logFile, FileMode.Append, FileAccess.Write, FileShare.None, bufferSize: 4096, useAsync: true)) {
                await sourceStream.WriteAsync(encodedText, 0, encodedText.Length);
            };
            ReportProgress($"Logfile saved to {logFile}");
        }

        /// <summary>
        /// Helper method to track logging for all of the processes 
        /// </summary>
        /// <param name="message"></param>
        public static void ReportProgress(string message)
        {
            var now = DateTime.Now.ToString(DT_FORMAT);
            message = $"{now}: {message}";
            Logger.AppendLine(message);
            Console.WriteLine(message);
        }
    }
}