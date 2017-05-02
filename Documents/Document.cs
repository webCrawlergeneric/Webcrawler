using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Documents
{
    class Document
    {
        static Document DocumentObj = null;
        private Document()
        {

        }

        public  static Document getDocumentObj()
        {
            if(DocumentObj==null)
            {
                Document obj = new Document();
                return obj;
            }
            else
            {
                return DocumentObj;
            }
        }
        public string Legalname { get; set; }
        public string AccountNumber { get; set; }
        public string Country { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Province { get; set; }
        public string Postalcode { get; set; }
        public string Phonenumber { get; set; }
        public string Ext {get; set; }
        public string Emailaddress { get; set; }
        public string WorkSafeBC_Legalname { get; set; }
        public string ClientAccoutnumber { get; set; }
        public string Legalname_Tradename { get; set; }
        public string ClientCode { get; set; }
        public string Status { get; set; }
        public string ClientCode_ClientName { get; set; }
        public string ClearanceDate { get; set; }


    }
    class ComboItem
    {
        public int ID { get; set; }
        public string Text { get; set; }
    }
}
