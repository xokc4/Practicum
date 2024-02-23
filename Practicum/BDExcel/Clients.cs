using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practicum.BDExcel
{
    /// <summary>
    ///  объект для клиента
    /// </summary>
    internal class Clients
    {
        public Clients(int clientCode, string organizationName, string address, string contactPerson)
        {
            ClientCode = clientCode;
            OrganizationName = organizationName;
            Address = address;
            ContactPerson = contactPerson;
        }

       public int ClientCode {  get; set; }// код клиента
       public string OrganizationName {  get; set; }// имя организации
      public  string Address {  get; set; }// адресс
      public  string ContactPerson {  get; set; }// имя клиента
    }
}
