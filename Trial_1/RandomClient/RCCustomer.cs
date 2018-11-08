using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trial_1.RandomClient
{
    class RCCustomer
    {
        private const int maxStatementLines = 15;

        public string CustomersFirstName { get; set; }
        public string CustomersMiddleName { get; set; }
        public string CustomersLastName { get; set; }
        public DateTime BillDate { get; set; }
        public int CustomerID { get; set; }
        public string AccountNumber { get; set; }
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string MailingCity { get; set; }
        public string MailingState { get; set; }
        public string MailingZipCode { get; set; }
        public string Subtotal { get; set; }
        public string TaxDue { get; set; }
        public string TotalDue { get; set; }

        public string IMBarcode { get; set; }
        public int SortPosition { get; set; }
        public int TrayNumber { get; set; }
        public int PageNumber => CustomerStatementSeparated.Count();

        public IEnumerable<List<RCCustomerStatement>> CustomerStatementSeparated { get; private set; }

        public RCCustomer(int aID)
        {
            CustomerID = aID;
        }

        private void pageSeparater(IEnumerable<RCCustomerStatement> aStatement)
        {
            List<List<RCCustomerStatement>> tempResult = new List<List<RCCustomerStatement>>();

            List<RCCustomerStatement> chunck = new List<RCCustomerStatement>();
            tempResult.Add(chunck);
            int pageLine = 0;
            for (int i = 0; i < aStatement.Count(); i++)
            {
                RCCustomerStatement currentStatement = aStatement.ElementAt(i);
                pageLine += currentStatement.DescriptionLine;

                if (pageLine > maxStatementLines)
                {
                    chunck = new List<RCCustomerStatement>();
                    chunck.Add(currentStatement);
                    tempResult.Add(chunck);
                    pageLine = currentStatement.DescriptionLine;
                }
                else
                {
                    chunck.Add(currentStatement);
                }
            }

            CustomerStatementSeparated = tempResult;
        }

        public void setDescription(IEnumerable<RCCustomerStatement> aStatement)
        {
            pageSeparater(aStatement);
        }

        public void UpdateCustomer(RCCustomer aCustomer)
        {
            if (aCustomer.CustomerID != CustomerID || AccountNumber != aCustomer.AccountNumber) throw new ArgumentException("Error");

            CustomersFirstName = aCustomer.CustomersFirstName;
            CustomersLastName = aCustomer.CustomersLastName;
            CustomersMiddleName = aCustomer.CustomersMiddleName;
            AddressLine1 = aCustomer.AddressLine1;
            AddressLine2 = aCustomer.AddressLine2;
            MailingCity = aCustomer.MailingCity;
            MailingState = aCustomer.MailingState;
            MailingZipCode = aCustomer.MailingZipCode;

            IMBarcode = aCustomer.IMBarcode;
            SortPosition = aCustomer.SortPosition;
            TrayNumber = aCustomer.TrayNumber;
        }

        public override string ToString()
        {
            string str = "";

            str += "CustomersFirstName-" + CustomersFirstName + ",";
            str += "CustomersMiddleName-" + CustomersMiddleName + ",";
            str += "CustomersLastName-" + CustomersLastName + ",";
            str += "BillDate-" + BillDate + ",";
            str += "AccountNumber-" + AccountNumber + ",";
            str += "AddressLine1-" + AddressLine1 + ",";
            str += "AddressLine2-" + AddressLine2 + ",";
            str += "MailCity-" + MailingCity + ",";
            str += "MailState-" + MailingState + ",";
            str += "MailZip-" + MailingZipCode + ",";
            str += "NumberOfStatement-" + CustomerStatementSeparated.Sum(s => s.Count) + ",";

            str += "IMBarcode-" + IMBarcode + ",";
            str += "SortPosition-" + SortPosition + ",";
            str += "TrayNumber-" + TrayNumber + ",";

            str += "PatientID-" + CustomerID + ",";
            return str;
        }
    }
}
