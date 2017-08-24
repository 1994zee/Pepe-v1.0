using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pepe.Model
{
    class Record
    {
        public string caseNo { get; set; }
        public string billDate { get; set; }
        public string billEventCode { get; set; }
        public string billRates { get; set; }
        public string billUnits { get; set; }
        public string clientId { get; set; }
        public string location { get; set; }
        public string comment { get; set; }

        //defualt constructor
        public Record()
        {
            caseNo = billDate = billEventCode = billRates = billUnits = clientId = location = comment = null;
        }
        //overlaoded constructor
        public Record(string a, string b, string c, string d, string e, string f,string g, string h)
        {
            caseNo = a;
            billDate = b;
            billEventCode = c;
            billRates = d;
            billUnits = e;
            clientId = f;
            location = g;
            comment = h;
        }
        //print function of record
        public void printData()
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(caseNo + "\t" + billDate + "\t" + billEventCode + "\t" + billRates + "\t" + billUnits + "\t" + clientId + "\t" + location + "\t" + comment);
            Console.ForegroundColor = ConsoleColor.White;
        }
        public string toString()
        {
            return caseNo + "\t" + billDate + "\t" + billEventCode + "\t" + billRates + "\t" + billUnits + "\t" + clientId + "\t" + location + "\t" + comment;
        }
    }
}
