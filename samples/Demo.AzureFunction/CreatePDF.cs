using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using PnP.Core.Services;
using PnP.Core.Model.SharePoint;
using System.Linq;
using System.IO;
using Spire.Doc;
using PnP.Core.QueryModel;
using System;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.ObjectModel;

namespace Demo.AzureFunction
{
    public class CreatePDF
    {
        private readonly IPnPContextFactory pnpContextFactory;
        public CreatePDF(IPnPContextFactory pnpContextFactory)
        {
            this.pnpContextFactory = pnpContextFactory;
        }

        [FunctionName("CreatePDF")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req,
            ILogger log)
        {

            log.LogInformation("CreatePDF()");
            using (var pnpContext = await pnpContextFactory.CreateAsync("Default"))
            {
                log.LogInformation("Creating PDF");

                IList sharedDocuments = await pnpContext.Web.Lists.GetByServerRelativeUrlAsync("/sites/Dokstyring/ForenkletDokumenter", l => l.RootFolder);

                var sharedDocumentsFolder = await sharedDocuments.RootFolder.GetAsync(f => f.Files);

                var arbeidsDoklist = pnpContext.Web.Lists.GetByServerRelativeUrl("/sites/Dokstyring/ForenkletDokumenter", p => p.Title, p => p.Items,
                                                     p => p.Fields.QueryProperties(p => p.InternalName,
                                                                                   p => p.FieldTypeKind,
                                                                                   p => p.TypeAsString,
                                                                                   p => p.Title));
                IFolder publiserteDokLib = await pnpContext.Web.GetFolderByServerRelativeUrlAsync($"{pnpContext.Uri.PathAndQuery}/ForenkletPubliserteDokumenter");
                IList publiserteDokList = await pnpContext.Web.Lists.GetByServerRelativeUrlAsync($"{pnpContext.Uri.PathAndQuery}/ForenkletPubliserteDokumenter");

                foreach (var listItem in arbeidsDoklist.Items.AsRequested())
                {
                    
                     try
                    {
                        var doc = await listItem.File.GetAsync();
                        string documentUrl = $"{pnpContext.Uri.PathAndQuery}/ForenkletDokumenter/{doc.Name}";
                        IFile workDocument = await pnpContext.Web.GetFileByServerRelativeUrlAsync(documentUrl);
                        Stream downloadedContentStream = await workDocument.GetContentAsync();
                        Document publishedDocument = new Document(downloadedContentStream, FileFormat.Auto);

                        var pdfName = $"{doc.Name}";
                        pdfName = pdfName.Replace("docx", "pdf");
                        //TODO: Save to stream and upload to sharepoint
                        publishedDocument.SaveToFile(pdfName, FileFormat.PDF);

                        IFile addedFile = await publiserteDokLib.Files.AddAsync(pdfName, System.IO.File.OpenRead($".{Path.DirectorySeparatorChar}{pdfName}"), true);
                        await addedFile.ListItemAllFields.LoadAsync();
                        await SetValuesToSingleTaxField(arbeidsDoklist, addedFile, listItem, "dsHovedprosess");
                        await SetValuesToMultiTaxField(arbeidsDoklist, addedFile, listItem, "dsDelprosess");
                        await SetValuesToMultiTaxField(arbeidsDoklist, addedFile, listItem, "dsFag");
                        //addedFile.ListItemAllFields["Title"] = listItem["Title"];
                        await addedFile.ListItemAllFields.UpdateAsync();

                    }
                    catch (System.Exception ex)
                    {
                        //return new BadRequestResult();
                        throw ex;
                    }
                    
                  
                    

                }



                return new OkObjectResult("OK");
            }
        }

        private static async Task SetValuesToSingleTaxField(IList arbeidsDoklist, IFile addedFile, IListItem listItem, string fieldInternalName)
        {
            var fieldValue = (listItem[fieldInternalName] as IFieldTaxonomyValue);
            IField field1 = await arbeidsDoklist.Fields.Where(f => f.InternalName == fieldInternalName).FirstOrDefaultAsync();
            if (fieldValue != null)
            {
                var val = field1.NewFieldTaxonomyValue(fieldValue.TermId, fieldValue.Label);
                addedFile.ListItemAllFields[fieldInternalName] = val;

            }
        }

        private static async Task SetValuesToMultiTaxField(IList arbeidsDoklist, IFile addedFile, IListItem listItem, string fieldInternalName)
        {
            var fieldValues = (listItem[fieldInternalName] as IFieldValueCollection)?.Values;
            IField field2 = await arbeidsDoklist.Fields.Where(f => f.InternalName == fieldInternalName).FirstOrDefaultAsync();
            if (fieldValues != null)
            {
                var taxonomyCollection = field2.NewFieldValueCollection();
                foreach (IFieldTaxonomyValue taxField in fieldValues)
                {
                    taxonomyCollection.Values.Add(field2.NewFieldTaxonomyValue(taxField.TermId, taxField.Label));
                }
                addedFile.ListItemAllFields[fieldInternalName] = taxonomyCollection;
            }
        }
    }
}
