using DocuWare.Platform.ServerClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;

namespace DocuWarePlatform.NETClient
{
    class PlatformClient
    {
        HttpClientHandler clientHandler;
        ServiceConnection connector;
        Organization org;

        public PlatformClient(string serverUrl, string organizationName, string userName, string userPassword)
        {

            this.connector = ServiceConnection.Create(new System.Uri(String.Format("{0}/docuware/platform", serverUrl)),
                                                      userName: userName,
                                                      password: userPassword,
                                                      organization: organizationName);

            this.org = this.connector.Organizations[0];
        }

        public void CloseConnection()
        {
            this.connector.Disconnect();
        }

        public IEnumerable<FileCabinet> GetAllFileCabinetsUserHasAccessTo(bool isBasket)
        {
            return (from fileCabinet in this.org.GetFileCabinetsFromFilecabinetsRelation().FileCabinet
                    where fileCabinet.IsBasket == isBasket
                    select fileCabinet);
        }

        public FileCabinet GetFileCabinet(string fileCabinetName, bool isBasket)
        {
            return (from fileCabinet in GetAllFileCabinetsUserHasAccessTo(isBasket)
                    where fileCabinet.Name.Equals(fileCabinetName)
                    select fileCabinet).SingleOrDefault();
        }

        public int GetTotalAmountOfDocuments(FileCabinet fileCabinet)
        {
            var searchDialog = getDefaultSearchDialog(fileCabinet);

            return searchDialog.GetCountResultFromCountRelation().Group.First().Count;
        }

        // This is the most performant way to get all documents at once.
        public List<Document> GetAllDocuments(FileCabinet fileCabinet)
        {
            // Check optional parameters of the method GetFromDocumentsForDocumentsQueryResultAsync.
            // Using them you can specify:
            //    * fields to retrieve
            //    * sort order
            //    * query
            //    * start index (which document do you want to start with)
            //    * max. amount of documents to retrieve
            return this.connector.GetFromDocumentsForDocumentsQueryResultAsync(fileCabinet.Id).Result.Content.Items;
        }

        /// <summary>
        /// The documents are returned in pages of the specified size.
        /// </summary>
        /// <param name="fileCabinet"> File cabinet to get documents from. </param>
        /// <param name="start"> The number of documents to be skipped, that is, the result list does not contain the first start documents. </param>
        /// <param name="maxCount"> The maximum number of items per result page. The server returns at most maxCount items per page. The actual number of items returned can be smaller.</param>
        /// <returns></returns>
        public DocumentsQueryResult GetDocumentsUsingPaging(FileCabinet fileCabinet, int start = 0, int maxCount = 3)
        {
            return this.connector.GetFromDocumentsForDocumentsQueryResultAsync(fileCabinet.Id, start: start, count: maxCount).Result.Content;
        }

        public List<Document> GetDocumentsByQuery(FileCabinet fileCabinet, DialogExpression query)
        {
            var searchDialog = getDefaultSearchDialog(fileCabinet);

            return runQueryForDocuments(searchDialog, query).Items;
        }

        public void DownloadDocumentThumbnail(Document document)
        {
            var thumbnail = document.GetStreamFromThumbnailRelation();
            using (var thumbnailFile = File.Create(@"C:\Temp\MyThumbnail.png"))
            {
                thumbnail.CopyTo(thumbnailFile);
            }
        }

        public void DeleteDocuments(List<Document> documentsToDelete)
        {
            foreach (var document in documentsToDelete)
                document.DeleteSelfRelation();
        }

        public DocumentsQueryResult MoveDocumentFromFileCabinetToBasketDropIndexValues(Document document, string fileCabinetId, FileCabinet basket)
        {
            var sourceDocument = new Document
            {
                Id = document.Id,
                // Needed in order to preserve document name.
                // (other index values will get lost)
                Fields = new List<DocumentIndexField> { DocumentIndexField.Create("DWWBDOCNAME", document.Title) }
            };

            var transferInfo = new DocumentsTransferInfo()
            {
                Documents = new List<Document>() { sourceDocument },
                KeepSource = false,    // the document will be moved, NOT copied
                SourceFileCabinetId = fileCabinetId
            };

            //
            // All index values, that were set in file cabinet, will get lost here!
            //
            return basket.PostToTransferRelationForDocumentsQueryResult(transferInfo);
        }

        // These method preserves index values when moving documment form file cabinet to a basket
        public DocumentsQueryResult MoveDocumentFromFileCabinetToBasket(int documentId, string fileCabinetId, FileCabinet basket)
        {
            var transferInfo = new FileCabinetTransferInfo()
            {
                KeepSource = false,   // the document will be moved, NOT copied
                SourceDocId = new List<int> { documentId },
                SourceFileCabinetId = fileCabinetId
            };

            return basket.PostToTransferRelationForDocumentsQueryResult(transferInfo);
        }

        public DocumentsQueryResult StoreDocumentFromBasketToFileCabinet(int documentId, string basketId, FileCabinet fileCabinet, List<DocumentIndexField> indexValues, bool keepDocumentInBasket = false)
        {
            var sourceDocument = new Document
            {
                Id = documentId,
                Fields = indexValues
            };

            var transferInfo = new DocumentsTransferInfo()
            {
                Documents = new List<Document>() { sourceDocument },
                KeepSource = keepDocumentInBasket,
                SourceFileCabinetId = basketId
            };

            return fileCabinet.PostToTransferRelationForDocumentsQueryResult(transferInfo);
        }

        public DocumentsQueryResult StoreDocumentFromBasketToFileCabinetUsingIntellixHints(int documentId, string basketId, FileCabinet fileCabinet)
        {
            var transferInfo = new FileCabinetTransferInfo()
            {
                KeepSource = false,
                SourceDocId = new List<int> { documentId },
                SourceFileCabinetId = basketId,
                FillIntellix = true
            };

            return fileCabinet.PostToTransferRelationForDocumentsQueryResult(transferInfo);
        }

        public DocumentIndexFields ChangeIndexValues(Document document, List<DocumentIndexField> indexValues)
        {
            var fields = new DocumentIndexFields()
            {
                Field = indexValues
            };

            return document.PutToFieldsRelationForDocumentIndexFields(fields);
        }

        /// <summary>
        /// Runs the query and change index values for each document found.
        /// </summary>
        /// <param name="fileCabinet"> File cabinet to search documents in. </param>
        /// <param name="query"> Query specifying the documents. </param>
        /// <param name="indexValues"> Index values to change. </param>
        /// <returns></returns>
        public List<BatchUpdateResultItem> ChangeIndexValuesInBatch(FileCabinet fileCabinet, DialogExpression query, List<DocumentIndexField> indexValues)
        {
            var searchDialog = getDefaultSearchDialog(fileCabinet);
            var storeDialog = getDefaultStoreDialog(fileCabinet);

            var queryResult = runQueryForDocuments(searchDialog, query);

            var batchUpdateData = new BatchUpdateProcessData()
            {
                BreakOnError = false,
                StoreDialogId = storeDialog.Id,
                Field = indexValues
            };

            return queryResult.PostToBatchUpdateRelationForBatchUpdateIndexFieldsResult(batchUpdateData).Item;
        }

        public Document ClipDocuments(List<int> documentIds, FileCabinet destination)
        {
            Document mergedDocument = destination.PutToContentMergeOperationRelationForDocument
                (
                    new ContentMergeOperationInfo()
                    {
                        Documents = documentIds,
                        Operation = ContentMergeOperation.Clip,
                        Force = true
                    }
                );
            return mergedDocument;
        }

        public DocumentsQueryResult SplitDocument(Document document, List<int> pages, FileCabinet destination)
        {
            if (document.ContentDivideOperationRelationLink == null)
                document = document.GetDocumentFromSelfRelation();

            DocumentsQueryResult result =
                document.PutToContentDivideOperationRelationForDocumentsQueryResult
                (
                    new ContentDivideOperationInfo()
                    {
                        Force = true,
                        Operation = ContentDivideOperation.Split,
                        Pages = pages,
                        ResultNames = new List<string> { "split document" }
                    }
                );
            return result;
        }

        public Document StapleDocuments(List<int> documentIds, FileCabinet destination)
        {
            Document mergedDocument = destination.PutToContentMergeOperationRelationForDocument
                (
                    new ContentMergeOperationInfo()
                    {
                        Documents = documentIds,
                        Operation = ContentMergeOperation.Staple,
                        Force = true
                    }
                );
            return mergedDocument;
        }

        /// <summary>
        /// This token will allows you to login with the same user credentials later.
        /// </summary>
        /// <param name="lifetime"> Defines the time spane after that the token will expire. </param>
        /// <returns></returns>
        public string GetMultiusageToken(TimeSpan lifetime)
        {
            return this.org.PostToLoginTokenRelationForString(
                        new TokenDescription()
                        {
                            TargetProducts = new List<DWProductTypes> { DWProductTypes.PlatformService },
                            Usage = TokenUsage.Multi,
                            Lifetime = lifetime.ToString()
                        }
                    );
        }

        private Dialog getDefaultSearchDialog(FileCabinet fileCabinet)
        {
            return fileCabinet.GetDialogInfosFromSearchesRelation().Dialog.Where(dlg => dlg.IsDefault == !fileCabinet.IsBasket).FirstOrDefault().GetDialogFromSelfRelation();
        }

        private Dialog getDefaultStoreDialog(FileCabinet fileCabinet)
        {
            return fileCabinet.GetDialogInfosFromStoresRelation().Dialog.Where(dlg => dlg.IsDefault == true).FirstOrDefault().GetDialogFromSelfRelation();
        }

        private DocumentsQueryResult runQueryForDocuments(Dialog dialog, DialogExpression query)
        {
            return dialog.Query.PostToDialogExpressionRelationForDocumentsQueryResult(query);
        }

        /// <summary>
        /// This is an alternative implemantation creating connection if you want to use persisted cookies.
        /// </summary>
        private void connectToPlatformServiceUsingCookies(string serverUrl, string organizationName, string userName, string userPassword)
        {
            this.clientHandler = new HttpClientHandler()
            {
                CookieContainer = getPersistetCookies(),
                AutomaticDecompression = System.Net.DecompressionMethods.GZip,
                AllowAutoRedirect = true,
                UseCookies = true
            };

            this.connector = ServiceConnection.Create(new System.Uri(String.Format("{0}/docuware/platform", serverUrl)),
                                           userName: userName,
                                           password: userPassword,
                                           organization: organizationName,
                                           httpClientHandler: this.clientHandler);

            persistCookies(clientHandler.CookieContainer);
        }

        /// <summary>
        /// Gets persisted cookies.
        /// </summary>
        /// <returns> Persisted cookies if there are any, otherwise an empty CookieContainer. </returns>
        private System.Net.CookieContainer getPersistetCookies()
        {
            var cookies = new System.Net.CookieContainer();

            // Here should be your implementation that fill persisted cookies
            // into the CookieContainer

            return cookies;
        }

        private void persistCookies(System.Net.CookieContainer cookies)
        {
            // Here should be your implementation that persists cookies.
        }
    }
}