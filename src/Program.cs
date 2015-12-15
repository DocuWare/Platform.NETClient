using DocuWare.Platform.ServerClient;
using System.Collections.Generic;
using System.Linq;

namespace DocuWarePlatform.NETClient
{
    class Program
    {
        #region PBRDTEST
        const string DocuWareServerUrl = @"https://pbrdtest.docuware.com";
        const string userName = "User";
        const string userPassword = "user";
        const string organizationName = "Peters Engineering";
        const string basketName = "Inbox";
        const string fileCabinetName = "Platform Workshop";
        const string anotherFileCabinetName = "Document Pool";
        #endregion

        static PlatformClient platformClient;

        static void Main(string[] args)
        {
            platformClient = new PlatformClient(DocuWareServerUrl, organizationName, userName, userPassword);

            var fileCabinet = platformClient.GetFileCabinet(fileCabinetName);
            // If you check fileCabinet.Fields property now, it will be "null".
            // That's because only part of the information about a file cabinet is provided at this point.
            // If you want to have the whole data available, run this call:
            fileCabinet = fileCabinet.GetFileCabinetFromSelfRelation();

            platformClient.CloseConnection();
        }

        private void DoSomethingWithDocumentUsingPaging(string targetName, bool isDocumentTray)
        {
            DocumentsQueryResult queryResult;
            List<Document> documents;
            int startPosition = 0;

            // Go through all pages
            do
            {
                // Get documents
                queryResult = platformClient.GetDocumentsUsingPaging(targetName, isDocumentTray, start: startPosition);
                documents = queryResult.Items;
                startPosition += documents.Count;

                // Do something with these documents
                // ...

            } while (queryResult.NextRelationLink != null);
        }

        // ToDo: try using properties of DialogExpressionCondition instead of calling Create method.
        private List<Document> GetDocumentsThatMeetSearchCriteriaSpecified(string targetName, bool isDocumentTray)
        {
            var query = new DialogExpression()
            {
                Operation = DialogExpressionOperation.And,
                Condition = new List<DialogExpressionCondition>()
                {
                    DialogExpressionCondition.Create(fieldName: "DWSTOREUSER", value: userName),
                    DialogExpressionCondition.Create(fieldName: "TOTAL_EFFORT", valueFrom: "10", valueTo: "50"),
                    DialogExpressionCondition.Create(fieldName: "KEYWORD", value: "five")
                },
                SortOrder = new List<SortedField>
                {
                    SortedField.Create("DWWBDOCNAME", SortDirection.Desc)
                }
            };

            return platformClient.GetDocumentsByQuery(targetName, isDocumentTray, query);
        }

        private Document GetDocumentById(string targetName, bool isDocumentTray, int documentId)
        {
            var query = new DialogExpression()
            {
                Operation = DialogExpressionOperation.And,
                Condition = new List<DialogExpressionCondition>()
                {
                    DialogExpressionCondition.Create(fieldName: "DWDOCID", value: documentId.ToString())
                },
            };

            return platformClient.GetDocumentsByQuery(targetName, isDocumentTray, query).FirstOrDefault();
        }

        private List<Document> SearchForDocumentsInSeveralFileCabinets(string firstFileCabinetName, string secondFileCabinetName)
        {
            var query = new DialogExpression()
            {
                Operation = DialogExpressionOperation.Or,
                Condition = new List<DialogExpressionCondition>()
                {
                    DialogExpressionCondition.Create(fieldName: "DWSTOREUSER", value: userName)
                },
                SortOrder = new List<SortedField>
                {
                    SortedField.Create("DWWBDOCNAME", SortDirection.Desc)
                },
                AdditionalCabinets = new List<string>() { platformClient.GetFileCabinet(secondFileCabinetName).Id }
            };

            return platformClient.GetDocumentsByQuery(firstFileCabinetName, isDocumentTray: false, query: query);
        }

        private void StoreDocumentLocatedInDocumentTrayIntoFileCabinet(string documentTrayName, string fileCabinetName, int documentId)
        {
            var document = GetDocumentById(documentTrayName, true, documentId);
            var documentTray = platformClient.GetDocumentTray(documentTrayName);
            var fileCabinet = platformClient.GetFileCabinet(fileCabinetName);
            var indexValues = new List<DocumentIndexField>
            {
                DocumentIndexField.Create("COMPANY", "SAGE"),
                DocumentIndexField.Create("CONTACT", "Elena"),
                DocumentIndexField.Create("SUBJECT", "Workshop"),
                DocumentIndexField.Create("DOCTYPE", "Example"),
                DocumentIndexField.Create("DATE", System.DateTime.Now),
                DocumentIndexField.Create("STATUS", "Done")
            };
            var queryResult = platformClient.StoreDocumentFromBasketToFileCabinet(document, documentTray, fileCabinet, indexValues);
        }

        private void StoreDocumentLocatedInDocumentTrayIntoFileCabinetUsingIntellixHints(string documentTrayName, string fileCabinetName, int documentId)
        {
            DocumentsQueryResult queryResult;
            var documentTray = platformClient.GetDocumentTray(documentTrayName);
            var fileCabinet = platformClient.GetFileCabinet(fileCabinetName);

            var document = GetDocumentById(documentTrayName, true, documentId);

            switch (document.IntellixTrust)
            {
                case IntellixTrust.Red:
                case IntellixTrust.Yellow:

                    // do you still want to use Intellix hints?
                    break;
                case IntellixTrust.Green:

                    queryResult = platformClient.StoreDocumentFromBasketToFileCabinetUsingIntellixHints(document, documentTray, fileCabinet);

                    break;
                case IntellixTrust.None:
                case IntellixTrust.InProgress:
                case IntellixTrust.Failed:

                    // Here it makes no sense to store this document using Intellix hints
                    break;
            }
        }

        private DocumentIndexFields ChangeIndexValuesForSingleDocument(Document document)
        {
            var indexValues = new List<DocumentIndexField>
            {
                DocumentIndexField.Create("DOCTYPE", "Example"),
                DocumentIndexField.Create("DATE", System.DateTime.Now),
                DocumentIndexField.Create("STATUS", "Changed")
            };

            return platformClient.ChangeIndexValues(document, indexValues);
        }

        private List<BatchUpdateResultItem> ChangeIndexValuesForSeveralDocumentsAtOnce(string fileCabinetName)
        {
            // File cabinet documents are stored in.
            var fileCabinet = platformClient.GetFileCabinet(fileCabinetName);

            var query = new DialogExpression()
            {
                Operation = DialogExpressionOperation.And,
                Condition = new List<DialogExpressionCondition>()
                {
                    DialogExpressionCondition.Create(fieldName: "DWSTOREUSER", value: userName),
                    DialogExpressionCondition.Create(fieldName: "TOTAL_EFFORT", valueFrom: "10", valueTo: "50")
                }
            };

            var indexValues = new List<DocumentIndexField>
            {
                DocumentIndexField.Create("STATUS", "Changed in batch"),
                DocumentIndexField.Create("DATE", System.DateTime.Now),
            };

            // Check which data the return value of ChangeIndexValuesInBatch() contains!
            // It could be useful for your implementation; especially the property ErrorMessage.
            return platformClient.ChangeIndexValuesInBatch(fileCabinet, query, indexValues);
        }

        private void SplitDocument (string documentTrayName, int documentId)
        {
            // Basket the document is currently stored into.
            var documentTray = platformClient.GetDocumentTray(documentTrayName);
            var documentToSplit = GetDocumentById(documentTrayName, true, documentId);

            platformClient.SplitDocument(documentToSplit, new List<int> { 2 }, documentTray);
        }

        private Document MergeTwoDocuments(string documentTrayName, int documentIdFirst, int documentIdSecond)
        {
            // Basket the document is currently stored into.
            var documentTray = platformClient.GetDocumentTray(documentTrayName);

            return platformClient.ClipDocuments(new List<int> { documentIdFirst, documentIdSecond}, documentTray);
        }

        private void DownloadDocumentThumbnail(Document document, string thumbmailFilePath)
        {
            var thumbnail = document.GetStreamFromThumbnailRelation();
            using (var thumbnailFile = System.IO.File.Create(thumbmailFilePath))
            {
                thumbnail.CopyTo(thumbnailFile);
            }
        }

        private void DeleteDocuments(List<Document> documentsToDelete)
        {
            foreach (var document in documentsToDelete)
                document.DeleteSelfRelation();
        }

    }
}