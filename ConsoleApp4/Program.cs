using System;
using Interop.QBFC16;

namespace ConsoleApp4
{
    public class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.DoInvoiceAdd();
        }

        public void DoInvoiceAdd()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                // Create the session Manager object
                sessionManager = new QBSessionManager();

                // Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildInvoiceAddRq(requestMsgSet);

                // Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                // Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);



                // End the session and close the connection to QuickBooks
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
                if (sessionBegun)
                {
                    sessionManager.EndSession();
                }
                if (connectionOpen)
                {
                    sessionManager.CloseConnection();
                }
            }
        }

        void BuildInvoiceAddRq(IMsgSetRequest requestMsgSet)
        {
            try
            {
                IInvoiceAdd invoiceAddRq = requestMsgSet.AppendInvoiceAddRq();


                invoiceAddRq.CustomerRef.FullName.SetValue("Thomas");
                invoiceAddRq.TxnDate.SetValue(DateTime.Today);
              
                string items = "APDEV";

                IORInvoiceLineAdd lineItem = invoiceAddRq.ORInvoiceLineAddList.Append();
                lineItem.InvoiceLineAdd.ItemRef.FullName.SetValue(items);
                lineItem.InvoiceLineAdd.Desc.SetValue($"Description for {items}");
                lineItem.InvoiceLineAdd.Quantity.SetValue(4);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{ex.Message}");
            }

        }
    }
}



