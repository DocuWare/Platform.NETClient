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

        #region DeploymentTest
        //const string DocuWareServerUrl = @"http://VSLM-503-003ecc2a-493e-4437-a51a-60aaa4e97bc9.DocuWare.AG";
        //const string userName = "admin";
        //const string userPassword = "admin";
        //const string organizationName = "Peters Engineering";
        //const string basketName = "Inbox";
        //const string fileCabinetName = "Document Pool";
        //const string anotherFileCabinetName = "Document Pool";
        #endregion

        static PlatformClient platformClient;
        static FileCabinet basket;
        static FileCabinet fileCabinet;

        static void Main(string[] args)
        {
            List<DocumentIndexField> indexValues;
            DialogExpression query;
            Document document;
            List<Document> documents;
            DocumentsQueryResult queryResult;
            DocumentIndexFields indexFields;

            platformClient = new PlatformClient(DocuWareServerUrl, organizationName, userName, userPassword);

            basket = platformClient.GetFileCabinet(basketName, isBasket: true);

            fileCabinet = platformClient.GetFileCabinet(fileCabinetName, isBasket: false);

            var searchDialog =
                fileCabinet.GetDialogInfosFromSearchesRelation()
                    .Dialog.Where(dlg => dlg.IsDefault == true)
                    .FirstOrDefault();

            // If you check fileCabinet.Fields property now, it will be "null".
            // That's because only part of the information about a file cabinet is provided at this point.
            // If you want to have the whole data available, run this call:
            //fileCabinet = fileCabinet.GetFileCabinetFromSelfRelation();

            #region Get total amount of documents stored in a file cabinet
            //            int totalAmount = platformClient.GetTotalAmountOfDocuments(fileCabinet);
            #endregion

            #region Get all documents stored in a fileCabinet
            //            documents = platformClient.GetAllDocuments(fileCabinet);
            #endregion

            #region Get documents using paging

            //int count = 0;
            //queryResult = platformClient.GetDocumentsUsingPaging(fileCabinet);
            //documents = queryResult.Items;
            //count += documents.Count;

            //// Do something with these documents

            //while (queryResult.NextRelationLink != null)
            //{
            //    // Get next documents
            //    queryResult = platformClient.GetDocumentsUsingPaging(fileCabinet, start: count);
            //    documents = queryResult.Items;
            //    count += documents.Count;

            //    // Do something with these documents
            //};

            #endregion

            #region Get documents that meet search criteria specified

            //query = new DialogExpression()
            //{
            //    Operation = DialogExpressionOperation.And,
            //    Condition = new List<DialogExpressionCondition>()
            //    {
            //        DialogExpressionCondition.Create(fieldName: "DWSTOREUSER", value: userName),
            //        DialogExpressionCondition.Create(fieldName: "TOTAL_EFFORT", valueFrom: "10", valueTo: "50"),
            //        DialogExpressionCondition.Create(fieldName: "KEYWORD", value: "five")     // searching in keyword field
            //    },
            //    SortOrder = new List<SortedField> 
            //    { 
            //        SortedField.Create("DWWBDOCNAME", SortDirection.Desc)
            //    }
            //};

            //documents = platformClient.GetDocumentsByQuery(fileCabinet, query);

            #endregion

            #region Search for document in several file cabinets

            //var anotherFileCabinet = platformClient.GetFileCabinet(anotherFileCabinetName, false);
            //query = new DialogExpression()
            //{
            //    Operation = DialogExpressionOperation.Or,
            //    Condition = new List<DialogExpressionCondition>()
            //    {
            //        DialogExpressionCondition.Create(fieldName: "DWSTOREUSER", value: userName)
            //    },
            //    SortOrder = new List<SortedField> 
            //    { 
            //        SortedField.Create("DWWBDOCNAME", SortDirection.Desc)
            //    },
            //    AdditionalCabinets = new List<string>() { anotherFileCabinet.Id }
            //};

            //documents = platformClient.GetDocumentsByQuery(fileCabinet, query);


            #endregion

            #region Delete documents

            //documents = new List<Document>() { getDocumentFromFileCabinet("17") };
            //platformClient.DeleteDocuments(documents);
            //documents = new List<Document>() { getDocumentFromBasket("821420802") };
            //platformClient.DeleteDocuments(documents);

            #endregion

            #region Move a particular document from file cabinet to a basket

            //document = getDocumentFromFileCabinet("17");
            //queryResult = platformClient.MoveDocumentFromFileCabinetToBasket(document.Id, fileCabinet.Id, basket);

            #endregion

            #region Store a particular document located in a basket to the file cabinet and set index values

            //document = getDocumentFromFileCabinet("821420802");
            //indexValues = new List<DocumentIndexField>
            //{
            //    DocumentIndexField.Create("COMPANY", "SAGE"),
            //    DocumentIndexField.Create("CONTACT", "Elena"),
            //    DocumentIndexField.Create("SUBJECT", "Workshop"),
            //    DocumentIndexField.Create("DOCTYPE", "Example"),
            //    DocumentIndexField.Create("DATE", System.DateTime.Now),
            //    DocumentIndexField.Create("STATUS", "Done")
            //};
            //queryResult = platformClient.StoreDocumentFromBasketToFileCabinet(document.Id,
            //                                                                  basket.Id,
            //                                                                  fileCabinet,
            //                                                                  indexValues);

            #endregion

            #region Store a particular document located in a basket to the file cabinet and set index values using Intellix hints

            //    document = getDocumentFromBasket("714495565");

            //    switch (document.IntellixTrust)
            //    {
            //        case IntellixTrust.Red:
            //        case IntellixTrust.Yellow:

            //            // do you still want to use Intellix hints?
            //            break;
            //        case IntellixTrust.Green:

            //            queryResult = platformClient.StoreDocumentFromBasketToFileCabinetUsingIntellixHints(document.Id, basket.Id, fileCabinet);

            //            break;
            //        case IntellixTrust.None:
            //        case IntellixTrust.InProgress:
            //        case IntellixTrust.Failed:

            //            // Here it makes no sence to store this document using Intellix hints
            //            break;
            //}
            #endregion

            #region Change index values for a single document

            //document = getDocumentFromFileCabinet("17");
            //indexValues = new List<DocumentIndexField>
            //{
            //    DocumentIndexField.Create("DOCTYPE", "Example"),
            //    DocumentIndexField.Create("DATE", System.DateTime.Now),
            //    DocumentIndexField.Create("STATUS", "Changed")
            //};

            //indexFields = platformClient.ChangeIndexValues(document, indexValues);

            #endregion

            #region Change index values for several documents at once.

            //query = new DialogExpression()
            //{
            //    Operation = DialogExpressionOperation.And,
            //    Condition = new List<DialogExpressionCondition>()
            //    {
            //        DialogExpressionCondition.Create(fieldName: "DWSTOREUSER", value: userName),
            //        DialogExpressionCondition.Create(fieldName: "TOTAL_EFFORT", valueFrom: "10", valueTo: "50")
            //    }
            //};

            //indexValues = new List<DocumentIndexField>
            //{
            //    DocumentIndexField.Create("STATUS", "Changed in batch"),
            //    DocumentIndexField.Create("DATE", System.DateTime.Now),
            //};

            //// Check which data result is containing, it could be useful for your implementation.
            //// Especially property ErrorMessage.
            //List<BatchUpdateResultItem> result = platformClient.ChangeIndexValuesInBatch(fileCabinet, query, indexValues);

            #endregion

            #region Merge two documents

            //List<int> docIds = platformClient.GetAllDocuments(basket).Select(doc => doc.Id).ToList();
            //Document mergedDocument = platformClient.ClipDocuments(docIds, basket);

            #endregion

            #region Split a document 

            //Document documentToSplit = getDocumentFromBasket("1232353453");
            //queryResult = platformClient.SplitDocument(documentToSplit, new List<int> { 2 }, basket);

            #endregion

            #region Get Thumbnail

            //document = platformClient.GetAllDocuments(basket).First();
            //platformClient.DownloadDocumentThumbnail(document);

            #endregion


            platformClient.CloseConnection();

        }

        static private Document getDocumentFromFileCabinet(string docId)
        {
            var query = new DialogExpression()
            {
                Operation = DialogExpressionOperation.And,
                Condition = new List<DialogExpressionCondition>()
                {
                    DialogExpressionCondition.Create(fieldName: "DWDOCID", value: docId)
                },
            };

            return platformClient.GetDocumentsByQuery(fileCabinet, query).FirstOrDefault();
        }

        static private Document getDocumentFromBasket(string docId)
        {
            var query = new DialogExpression()
            {
                Operation = DialogExpressionOperation.And,
                Condition = new List<DialogExpressionCondition>()
                {
                    DialogExpressionCondition.Create(fieldName: "DWDOCID", value: docId)
                },
            };

            return platformClient.GetDocumentsByQuery(basket, query).FirstOrDefault();
        }
    }
}