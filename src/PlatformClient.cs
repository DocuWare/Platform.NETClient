using DocuWare.Platform.ServerClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;

namespace DocuWarePlatform.NETClient
{
    /// <summary>
    /// This is a sample how a PlatformClient could be implemented.
    /// </summary>
    /// <remarks> This implementation hardly contains error handling, it is just supposed to be used if you are starting with you first PlatformClient and don't know where or how to begin. </remarks>
    class PlatformClient
    {
        HttpClientHandler clientHandler;
        ServiceConnection connector;
        Organization org;


        /// <summary> Constructor. </summary>
        /// <param name="serverUrl"> The URL of the server Platform is running on. </param>
        /// <param name="organizationName"> Name of the organization you want the client to be connected to.</param>
        /// <param name="userName"> User to use when connecting to the organization specified.</param>
        /// <param name="userPassword"> Password of the user specified. </param>
        /// <remarks>
        /// When this constructor is running a connection to DocuWare Platform will be created; this action potentially consumes a license.
        /// "Potentially" means that not every run of this constructor will consume a license, in many cases the license will be re-used.
        /// We are not going into the details here because the underlying implementation regarding license consumption can be changed in the future.
        /// </remarks>
        public PlatformClient(string serverUrl, string organizationName, string userName, string userPassword)
        {

            this.connector = ServiceConnection.Create(new System.Uri(String.Format("{0}/docuware/platform", serverUrl)),
                                                      userName: userName,
                                                      password: userPassword,
                                                      organization: organizationName);

            this.org = this.connector.Organizations[0];
        }


        /// <summary>
        /// This method should be called when you have completed the task(s) you needed the Platform client for and want the license to be released as soon as possible.
        /// </summary>
        /// <remarks>
        /// The license will NOT be released immediately if this method is called. It remains consumed for several minutes.
        /// It is NOT required to call this method since if the instance of PlatformClient gets destroyed the license will be released automatically,
        /// however, it will remains consumed longer.
        /// </remarks>
        public void CloseConnection()
        {
            this.connector.Disconnect();
        }


        /// <summary>
        /// Gives you access to all file cabinets the user that has been used when creating this instance of PlatformClient has access to.
        /// </summary>
        public IEnumerable<FileCabinet> GetAllFileCabinetsUserHasAccessTo()
        {
            return (from fileCabinet in this.org.GetFileCabinetsFromFilecabinetsRelation().FileCabinet
                    where fileCabinet.IsBasket == false
                    select fileCabinet);
        }


        /// <summary>
        /// Gives you access to document trays (baskets) the user that has been used when creating this instance of PlatformClient has access to.
        /// </summary>
        public IEnumerable<FileCabinet> GetAllDocumentTraysUserHasAccessTo()
        {
            return (from fileCabinet in this.org.GetFileCabinetsFromFilecabinetsRelation().FileCabinet
                    where fileCabinet.IsBasket == true
                    select fileCabinet);
        }


        /// <summary> Use it to access to a particular file cabinet. </summary>
        /// <param name="fileCabinetName"> Name of the file cabinet (case insensitive). </param>
        /// <returns> FileCabinet or null if no one found. </returns>
        public FileCabinet GetFileCabinet(string fileCabinetName)
        {
            return (from fileCabinet in GetAllFileCabinetsUserHasAccessTo()
                    where String.Compare( fileCabinet.Name, fileCabinetName, ignoreCase: true) == 0
                    select fileCabinet).SingleOrDefault();
        }


        /// <summary> Use it to access to a particular document tray (basket). </summary>
        /// <param name="documentTrayName"> Name of the file cabinet (case insensitive). </param>
        /// <returns> Document tray (basket) or null if no one found. </returns>
        public FileCabinet GetDocumentTray(string documentTrayName)
        {
            return (from documentTray in GetAllDocumentTraysUserHasAccessTo()
                    where String.Compare(documentTray.Name, documentTrayName, ignoreCase: true) == 0
                    select documentTray).SingleOrDefault();
        }


        /// <summary>
        /// This is the most effective way to get the total amount or documents in a file cabinet or document tray (basket).
        /// </summary>
        /// <param name="targetName"> The name if file cabinet or document tray (basket). </param>
        /// <param name="isDocumentTray"> Specifies if the target is a document tray (basket). </param>
        /// <returns> Total amount of documents. </returns>
        public int GetTotalAmountOfDocuments(string targetName, bool isDocumentTray)
        {
            var target = isDocumentTray ? GetDocumentTray(targetName) : GetFileCabinet(targetName);
            
            var searchDialog = getDefaultSearchDialog(target);

            return searchDialog.GetCountResultFromCountRelation().Group.First().Count;
        }


        /// <summary> This is the most effective way to get all documents at once. </summary>
        /// <param name="targetName"> The name if file cabinet or document tray (basket). </param>
        /// <param name="isDocumentTray"> Specifies if the target is a document tray (basket). </param>
        /// <returns> All documents currently stored in the location specified. </returns>
        public List<Document> GetAllDocuments(string targetName, bool isDocumentTray)
        {
            var target = isDocumentTray ? GetDocumentTray(targetName) : GetFileCabinet(targetName);

            // Check optional parameters of the method GetFromDocumentsForDocumentsQueryResultAsync.
            // Using them you can specify:
            //    * fields to retrieve
            //    * sort order
            //    * query
            //    * start index (which document do you want to start with)
            //    * max. amount of documents to retrieve
            return this.connector.GetFromDocumentsForDocumentsQueryResultAsync(target.Id, count: int.MaxValue).Result.Content.Items;
        }


        /// <summary> The documents are returned in pages of the size specified by maxCount. </summary>
        /// <param name="targetName"> The name if file cabinet or document tray (basket). </param>
        /// <param name="isDocumentTray"> Specifies if the target is a document tray (basket). </param>
        /// <param name="start"> The number of documents to be skipped, that is, the result list does not contain the first "start" documents. </param>
        /// <param name="maxCount"> The maximum number of items per result page. The server returns at most maxCount items per page. The actual number of items returned can be smaller.</param>
        /// <returns> TODO </returns>
        public DocumentsQueryResult GetDocumentsUsingPaging(string targetName, bool isDocumentTray, int start = 0, int maxCount = 3)
        {
            var target = isDocumentTray ? GetDocumentTray(targetName) : GetFileCabinet(targetName);

            return this.connector.GetFromDocumentsForDocumentsQueryResultAsync(target.Id, start: start, count: maxCount).Result.Content;
        }


        /// <summary> Use it if you are searching for documents. </summary>
        /// <param name="targetName"> The name if file cabinet or document tray (basket). </param>
        /// <param name="isDocumentTray"> Specifies if the target is a document tray (basket). </param>
        /// <param name="query"> Searching criteria. </param>
        /// <returns> List containing documents found. </returns>
        public List<Document> GetDocumentsByQuery(string targetName, bool isDocumentTray, DialogExpression query)
        {
            var target = isDocumentTray ? GetDocumentTray(targetName) : GetFileCabinet(targetName);
            var searchDialog = getDefaultSearchDialog(target);

            return runQueryForDocuments(searchDialog, query).Items;
        }


        /// <summary> Moves the document from file cabinet to document tray. </summary>
        /// <remarks> This implementation does not preserve index values when moving the document. </remarks>
        /// <param name="document"> Document to be moved. </param>
        /// <param name="fileCabinet"> File cabinet where document is currently located. </param>
        /// <param name="documentTray"> Document tray document has to be moved into. </param>
        /// <returns> TODO </returns>
        public DocumentsQueryResult MoveDocumentFromFileCabinetToBasketDropIndexValues(Document document, FileCabinet fileCabinet, FileCabinet documentTray)
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
                SourceFileCabinetId = fileCabinet.Id
            };

            //
            // All index values, that were set in file cabinet, will get lost here!
            //
            return documentTray.PostToTransferRelationForDocumentsQueryResult(transferInfo);
        }


        /// <summary> Moves the document from file cabinet to document tray. </summary>
        /// <remarks> This implementation preserves index values when moving the document. </remarks>
        /// <param name="document"> Document to be moved. </param>
        /// <param name="fileCabinet"> File cabinet where document is currently located. </param>
        /// <param name="documentTray"> Document tray document has to be moved into. </param>
        /// <returns> TODO </returns>
        public DocumentsQueryResult MoveDocumentFromFileCabinetToBasket(Document document, FileCabinet fileCabinet, FileCabinet documentTray)
        {
            var transferInfo = new FileCabinetTransferInfo()
            {
                KeepSource = false,   // the document will be moved, NOT copied
                SourceDocId = new List<int> { document.Id},
                SourceFileCabinetId = fileCabinet.Id
            };

            return documentTray.PostToTransferRelationForDocumentsQueryResult(transferInfo);
        }


        /// <summary> Shows how to store a document into file cabinet using index values provided. </summary>
        /// <param name="document"> Id of the document to be stored. </param>
        /// <param name="documentTray"> Document tray (basket) document is currently stored in. </param>
        /// <param name="fileCabinet"> File cabinet to store document into. </param>
        /// <param name="indexValues"> Index values to apply to the document when storing. </param>
        /// <param name="keepDocumentInDocumentTray"> Specifies if document should remains in document tray after storing. </param>
        /// <returns> TODO </returns>
        public DocumentsQueryResult StoreDocumentFromBasketToFileCabinet(Document document, FileCabinet documentTray, FileCabinet fileCabinet, List<DocumentIndexField> indexValues, bool keepDocumentInDocumentTray = false)
        {
            var sourceDocument = new Document
            {
                Id = document.Id,
                Fields = indexValues
            };

            var transferInfo = new DocumentsTransferInfo()
            {
                Documents = new List<Document>() { sourceDocument },
                KeepSource = keepDocumentInDocumentTray,
                SourceFileCabinetId = documentTray.Id
            };

            return fileCabinet.PostToTransferRelationForDocumentsQueryResult(transferInfo);
        }


        /// <summary> Shows how to store a document in a file cabinet using Intellix hints. </summary>
        /// <param name="document"> Id of the document to be stored. </param>
        /// <param name="documentTray"> Id of the document tray (basket) document is currently stored in. </param>
        /// <param name="fileCabinet"> File cabinet to store document into. </param>
        /// <returns> TODO </returns>
        public DocumentsQueryResult StoreDocumentFromBasketToFileCabinetUsingIntellixHints(Document document, FileCabinet documentTray, FileCabinet fileCabinet)
        {
            var transferInfo = new FileCabinetTransferInfo()
            {
                KeepSource = false,
                SourceDocId = new List<int> { document.Id },
                SourceFileCabinetId = documentTray.Id,
                FillIntellix = true
            };

            return fileCabinet.PostToTransferRelationForDocumentsQueryResult(transferInfo);
        }


        /// <summary> Shows how to change index values of a document. </summary>
        /// <param name="document"> Document who's index values should be changed. </param>
        /// <param name="indexValues"> New index values. </param>
        /// <returns> Document's index values after performing this operation. </returns>
        public DocumentIndexFields ChangeIndexValues(Document document, List<DocumentIndexField> indexValues)
        {
            var fields = new DocumentIndexFields()
            {
                Field = indexValues
            };

            return document.PutToFieldsRelationForDocumentIndexFields(fields);
        }


        /// <summary> Shows how to change index values for several documents at once. </summary>
        /// <remarks> Runs the query and change index values for each document found. </remarks>
        /// <param name="fileCabinet"> File cabinet to search documents in. TODO: also document tray?</param>
        /// <param name="query"> Query specifying documents who's index values should be changed. </param>
        /// <param name="indexValues"> Index values to change. </param>
        /// <returns> TODO </returns>
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


        /// <summary> Shows how to clip several documents. </summary>
        /// <param name="documentIds"> Documents to be clipped. </param>
        /// <param name="target"> File cabinet to where the single documents are currently stored. TODO: can you clip documents in basket? </param>
        /// <returns> Single document after clipping. </returns>
        public Document ClipDocuments(List<int> documentIds, FileCabinet target)
        {
            Document mergedDocument = target.PutToContentMergeOperationRelationForDocument
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


        /// <summary> Shows how to split document into several parts. </summary>
        /// <param name="document"> Document to be split. </param>
        /// <param name="pages"> TODO ??? </param>
        /// <param name="destination"> File cabinet the document is currently stored into.  TODO: can you split documents in basket??? </param>
        /// <returns> TODO </returns>
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


        /// <summary> Shows how to staple several documents. </summary>
        /// <param name="documents"> Documents to be stapled. </param>
        /// <param name="destination"> File cabinet the stapled document should be stored into.TODO: can you staple documents in basket??? </param>
        /// <returns> Stapled document. </returns>
        public Document StapleDocuments(List<Document> documents, FileCabinet destination)
        {
            Document mergedDocument = destination.PutToContentMergeOperationRelationForDocument
                (
                    new ContentMergeOperationInfo()
                    {
                        Documents = documents.Select(document => document.Id).ToList(),
                        Operation = ContentMergeOperation.Staple,
                        Force = true
                    }
                );
            return mergedDocument;
        }


        /// <summary> This token will allows you to login with the same user credentials later. </summary>
        /// <param name="lifetime"> Defines the time span after that the token will expire. </param>
        /// <returns> Token created. </returns>
        /// <remarks> TODO: what is the use case for that? </remarks>
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

        #region private methods

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

        #endregion


        #region Client that uses persisted cookies
        /// <summary> This is an alternative implementation creating connection if you want to use persisted cookies. </summary>
        /// <remarks> TODO </remarks>
        private void connectToPlatformServiceUsingCookies(string serverUrl, string organizationName, string userName, string userPassword)
        {
            this.clientHandler = new HttpClientHandler()
            {
                CookieContainer = getPersistedCookies(),
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

        /// <summary> Gets persisted cookies. </summary>
        /// <returns> Persisted cookies if there are any, otherwise an empty CookieContainer. </returns>
        private System.Net.CookieContainer getPersistedCookies()
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
        #endregion

    }
}